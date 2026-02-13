const functions = require("firebase-functions");
const admin = require("firebase-admin");
const cors = require("cors");
const Stripe = require("stripe");

// Initialize Firebase Admin
admin.initializeApp();
const db = admin.firestore();

// Initialize Stripe with secret key from Firebase config
const stripe = new Stripe(process.env.STRIPE_SECRET_KEY || functions.config().stripe?.secret_key);

// CORS middleware
const corsHandler = cors({ origin: true });

// =============================================================================
// LICENSE CHECK API
// Called by the Google Sheets add-on to verify user subscription
// =============================================================================
exports.checkLicense = functions.https.onRequest((req, res) => {
    corsHandler(req, res, async () => {
        try {
            // Only allow GET and POST
            if (req.method !== "GET" && req.method !== "POST") {
                return res.status(405).json({ error: "Method not allowed" });
            }

            // Get email from query params or body
            const email = req.query.email || req.body?.email;

            if (!email) {
                return res.status(400).json({
                    error: "Email is required",
                    allowed: false
                });
            }

            // Normalize email
            const normalizedEmail = email.toLowerCase().trim();
            const domain = normalizedEmail.split("@")[1];

            // Check for individual email license first
            const licenseRef = db.collection("licenses").doc(normalizedEmail);
            const licenseDoc = await licenseRef.get();

            // Check for domain-wide license (company purchase)
            const domainLicenseRef = db.collection("domain_licenses").doc(domain);
            const domainLicenseDoc = await domainLicenseRef.get();

            // Determine which license to use (individual takes priority for features)
            let license = null;
            let licenseType = null;
            let activeLicenseRef = null;

            if (licenseDoc.exists && isLicenseActive(licenseDoc.data())) {
                license = licenseDoc.data();
                licenseType = "individual";
                activeLicenseRef = licenseRef;
            } else if (domainLicenseDoc.exists && isLicenseActive(domainLicenseDoc.data())) {
                license = domainLicenseDoc.data();
                licenseType = "domain";
                activeLicenseRef = domainLicenseRef;
            }

            if (!license) {
                return res.status(200).json({
                    allowed: false,
                    plan: "free",
                    message: "No active subscription found"
                });
            }

            // Check expiration for non-recurring plans
            if (license.expiresAt && license.expiresAt.toDate() < new Date()) {
                await activeLicenseRef.update({ status: "expired" });
                return res.status(200).json({
                    allowed: false,
                    plan: "free",
                    message: "Subscription has expired"
                });
            }

            // Valid subscription!
            return res.status(200).json({
                allowed: true,
                plan: license.plan || "pro",
                status: license.status,
                licenseType: licenseType,
                domain: licenseType === "domain" ? domain : null,
                customerId: license.stripeCustomerId,
                expiresAt: license.currentPeriodEnd?.toDate()?.toISOString() || null,
                features: license.features || ["unlimited_syncs", "priority_support"],
                seats: license.seats || null,
                usedSeats: license.usedSeats || null
            });

        } catch (error) {
            console.error("Error checking license:", error);
            return res.status(500).json({
                error: "Internal server error",
                allowed: false
            });
        }
    });
});

// Helper function to check if a license is active
function isLicenseActive(license) {
    return license && (license.status === "active" || license.status === "trialing");
}

// =============================================================================
// STRIPE WEBHOOK HANDLER
// Processes Stripe events to manage licenses automatically
// =============================================================================
exports.stripeWebhook = functions.https.onRequest(async (req, res) => {
    const sig = req.headers["stripe-signature"];
    const webhookSecret = process.env.STRIPE_WEBHOOK_SECRET || functions.config().stripe?.webhook_secret;

    let event;

    try {
        // Verify webhook signature
        event = stripe.webhooks.constructEvent(
            req.rawBody,
            sig,
            webhookSecret
        );
    } catch (err) {
        console.error("Webhook signature verification failed:", err.message);
        return res.status(400).send(`Webhook Error: ${err.message}`);
    }

    console.log(`Processing Stripe event: ${event.type}`);

    try {
        switch (event.type) {
            // =========================
            // CHECKOUT COMPLETED
            // =========================
            case "checkout.session.completed": {
                const session = event.data.object;
                await handleCheckoutCompleted(session);
                break;
            }

            // =========================
            // SUBSCRIPTION CREATED
            // =========================
            case "customer.subscription.created": {
                const subscription = event.data.object;
                await handleSubscriptionCreated(subscription);
                break;
            }

            // =========================
            // SUBSCRIPTION UPDATED
            // =========================
            case "customer.subscription.updated": {
                const subscription = event.data.object;
                await handleSubscriptionUpdated(subscription);
                break;
            }

            // =========================
            // SUBSCRIPTION DELETED/CANCELED
            // =========================
            case "customer.subscription.deleted": {
                const subscription = event.data.object;
                await handleSubscriptionDeleted(subscription);
                break;
            }

            // =========================
            // INVOICE PAID (renewal)
            // =========================
            case "invoice.paid": {
                const invoice = event.data.object;
                await handleInvoicePaid(invoice);
                break;
            }

            // =========================
            // INVOICE PAYMENT FAILED
            // =========================
            case "invoice.payment_failed": {
                const invoice = event.data.object;
                await handleInvoicePaymentFailed(invoice);
                break;
            }

            default:
                console.log(`Unhandled event type: ${event.type}`);
        }

        res.status(200).json({ received: true });
    } catch (error) {
        console.error(`Error processing ${event.type}:`, error);
        res.status(500).json({ error: "Webhook handler failed" });
    }
});

// =============================================================================
// WEBHOOK HANDLERS
// =============================================================================

async function handleCheckoutCompleted(session) {
    const customerEmail = session.customer_email || session.customer_details?.email;

    if (!customerEmail) {
        console.error("No customer email in checkout session");
        return;
    }

    const normalizedEmail = customerEmail.toLowerCase().trim();
    const metadata = session.metadata || {};
    const licenseType = metadata.licenseType || "individual";
    const domain = metadata.domain;
    const seats = metadata.seats ? parseInt(metadata.seats) : null;

    // Get or create customer info
    const customer = await stripe.customers.retrieve(session.customer);

    const baseLicenseData = {
        stripeCustomerId: session.customer,
        stripeSubscriptionId: session.subscription,
        status: "active",
        plan: "pro",
        features: ["unlimited_syncs", "priority_support", "advanced_filters"],
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        checkoutSessionId: session.id,
        purchaserEmail: normalizedEmail
    };

    if (licenseType === "domain" && domain) {
        // Create domain-wide license
        await db.collection("domain_licenses").doc(domain).set({
            ...baseLicenseData,
            domain: domain,
            seats: seats,
            usedSeats: 0,
            plan: "enterprise", // Domain licenses get enterprise features
            features: getFeaturesForPlan("enterprise")
        }, { merge: true });

        console.log(`Domain license created for: ${domain} (purchased by: ${normalizedEmail})`);
    } else {
        // Create individual license
        await db.collection("licenses").doc(normalizedEmail).set({
            ...baseLicenseData,
            email: normalizedEmail
        }, { merge: true });

        console.log(`License created for: ${normalizedEmail}`);
    }
}

async function handleSubscriptionCreated(subscription) {
    const customer = await stripe.customers.retrieve(subscription.customer);
    const email = customer.email?.toLowerCase().trim();

    if (!email) {
        console.error("No email found for customer:", subscription.customer);
        return;
    }

    // Determine plan from price
    const plan = getPlanFromSubscription(subscription);

    await db.collection("licenses").doc(email).set({
        email: email,
        stripeCustomerId: subscription.customer,
        stripeSubscriptionId: subscription.id,
        status: subscription.status,
        plan: plan,
        features: getFeaturesForPlan(plan),
        currentPeriodStart: admin.firestore.Timestamp.fromDate(
            new Date(subscription.current_period_start * 1000)
        ),
        currentPeriodEnd: admin.firestore.Timestamp.fromDate(
            new Date(subscription.current_period_end * 1000)
        ),
        cancelAtPeriodEnd: subscription.cancel_at_period_end,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });

    console.log(`Subscription created for: ${email}, plan: ${plan}`);
}

async function handleSubscriptionUpdated(subscription) {
    const customer = await stripe.customers.retrieve(subscription.customer);
    const email = customer.email?.toLowerCase().trim();

    if (!email) return;

    const plan = getPlanFromSubscription(subscription);

    await db.collection("licenses").doc(email).update({
        status: subscription.status,
        plan: plan,
        features: getFeaturesForPlan(plan),
        currentPeriodStart: admin.firestore.Timestamp.fromDate(
            new Date(subscription.current_period_start * 1000)
        ),
        currentPeriodEnd: admin.firestore.Timestamp.fromDate(
            new Date(subscription.current_period_end * 1000)
        ),
        cancelAtPeriodEnd: subscription.cancel_at_period_end,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`Subscription updated for: ${email}, status: ${subscription.status}`);
}

async function handleSubscriptionDeleted(subscription) {
    const customer = await stripe.customers.retrieve(subscription.customer);
    const email = customer.email?.toLowerCase().trim();

    if (!email) return;

    await db.collection("licenses").doc(email).update({
        status: "canceled",
        canceledAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`Subscription canceled for: ${email}`);
}

async function handleInvoicePaid(invoice) {
    if (!invoice.subscription) return;

    const customer = await stripe.customers.retrieve(invoice.customer);
    const email = customer.email?.toLowerCase().trim();

    if (!email) return;

    // Refresh subscription status
    const subscription = await stripe.subscriptions.retrieve(invoice.subscription);

    await db.collection("licenses").doc(email).update({
        status: "active",
        currentPeriodEnd: admin.firestore.Timestamp.fromDate(
            new Date(subscription.current_period_end * 1000)
        ),
        lastPaymentAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`Invoice paid for: ${email}`);
}

async function handleInvoicePaymentFailed(invoice) {
    const customer = await stripe.customers.retrieve(invoice.customer);
    const email = customer.email?.toLowerCase().trim();

    if (!email) return;

    await db.collection("licenses").doc(email).update({
        status: "past_due",
        paymentFailedAt: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`Payment failed for: ${email}`);
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

function getPlanFromSubscription(subscription) {
    // Get the price ID from the subscription
    const priceId = subscription.items?.data?.[0]?.price?.id;

    // Map price IDs to plans (configure these in Stripe)
    const priceToPlan = {
        // Add your Stripe Price IDs here
        // "price_xxxxx": "pro",
        // "price_yyyyy": "enterprise",
    };

    return priceToPlan[priceId] || "pro";
}

function getFeaturesForPlan(plan) {
    const planFeatures = {
        free: [],
        pro: [
            "unlimited_syncs",
            "priority_support",
            "advanced_filters",
            "custom_fields"
        ],
        enterprise: [
            "unlimited_syncs",
            "priority_support",
            "advanced_filters",
            "custom_fields",
            "team_management",
            "api_access",
            "dedicated_support"
        ]
    };

    return planFeatures[plan] || planFeatures.pro;
}

// =============================================================================
// CREATE CHECKOUT SESSION
// For self-serve purchases from your website or add-on
// Supports both individual and domain-wide (company) purchases
// =============================================================================
exports.createCheckoutSession = functions.https.onRequest((req, res) => {
    corsHandler(req, res, async () => {
        try {
            if (req.method !== "POST") {
                return res.status(405).json({ error: "Method not allowed" });
            }

            const {
                email,
                priceId,
                successUrl,
                cancelUrl,
                licenseType = "individual", // "individual" or "domain"
                domain,                      // Required if licenseType is "domain"
                seats                        // Optional: number of seats for domain license
            } = req.body;

            if (!email || !priceId) {
                return res.status(400).json({
                    error: "Email and priceId are required"
                });
            }

            if (licenseType === "domain" && !domain) {
                return res.status(400).json({
                    error: "Domain is required for domain-wide licenses"
                });
            }

            // Check if customer already exists
            const customers = await stripe.customers.list({ email: email, limit: 1 });
            let customerId;

            if (customers.data.length > 0) {
                customerId = customers.data[0].id;
            }

            // Create Stripe Checkout Session
            const session = await stripe.checkout.sessions.create({
                customer: customerId,
                customer_email: customerId ? undefined : email,
                payment_method_types: ["card"],
                line_items: [
                    {
                        price: priceId,
                        quantity: seats || 1,
                    },
                ],
                mode: "subscription",
                success_url: successUrl || "https://your-domain.com/success?session_id={CHECKOUT_SESSION_ID}",
                cancel_url: cancelUrl || "https://your-domain.com/cancel",
                metadata: {
                    source: "jira_sync_addon",
                    licenseType: licenseType,
                    domain: licenseType === "domain" ? domain.toLowerCase().trim() : null,
                    seats: seats ? seats.toString() : null,
                    purchaserEmail: email.toLowerCase().trim()
                }
            });

            res.status(200).json({
                sessionId: session.id,
                url: session.url
            });

        } catch (error) {
            console.error("Error creating checkout session:", error);
            res.status(500).json({ error: error.message });
        }
    });
});

// =============================================================================
// CUSTOMER PORTAL
// Allow users to manage their subscription
// =============================================================================
exports.createPortalSession = functions.https.onRequest((req, res) => {
    corsHandler(req, res, async () => {
        try {
            if (req.method !== "POST") {
                return res.status(405).json({ error: "Method not allowed" });
            }

            const { email, returnUrl } = req.body;

            if (!email) {
                return res.status(400).json({ error: "Email is required" });
            }

            // Find customer by email
            const customers = await stripe.customers.list({
                email: email.toLowerCase().trim(),
                limit: 1
            });

            if (customers.data.length === 0) {
                return res.status(404).json({ error: "No subscription found for this email" });
            }

            // Create portal session
            const session = await stripe.billingPortal.sessions.create({
                customer: customers.data[0].id,
                return_url: returnUrl || "https://your-domain.com/account",
            });

            res.status(200).json({ url: session.url });

        } catch (error) {
            console.error("Error creating portal session:", error);
            res.status(500).json({ error: error.message });
        }
    });
});
