/**
 * Cloudflare Worker for Jira Proxy & License Management
 */

const ALLOWED_DOMAIN_SUFFIX = ".atlassian.net";

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);
        const path = url.pathname;

        // 1. Handle CORS preflight
        if (request.method === "OPTIONS") {
            return new Response(null, {
                headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Atlassian-Token, X-Proxy-Secret, X-Mock-Request",
                },
            });
        }

        // 2. Handle License API
        if (path === "/api/license/check") {
            return await handleLicenseCheck(request, env);
        }

        // 3. Handle Stripe Checkout
        if (path === "/api/stripe/create-session") {
            return await handleStripeSession(request, env);
        }

        // 4. Handle Stripe Webhook
        if (path === "/api/stripe/webhook") {
            return await handleStripeWebhook(request, env);
        }

        // 5. Debug Endpoint
        if (path === "/api/debug") {
            return new Response(JSON.stringify({
                hasKV: !!env.LICENSES,
                time: new Date().toISOString(),
                kvId: env.LICENSES ? "Present" : "Missing"
            }), { headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" } });
        }

        // 6. Build Sales CSV
        if (path === "/api/admin/sales-csv") {
            return await handleSalesExport(request, env);
        }

        // 6. Handle Jira Proxy
        if (url.searchParams.has("target")) {
            return await handleProxy(request, env);
        }

        // 7. Handle Static Assets / Landing Page
        return await handleStaticAssets(request, env);
    },
};

/**
 * Exports sales data from KV as a CSV
 */
async function handleSalesExport(request, env) {
    const secret = new URL(request.url).searchParams.get("secret");
    if (secret !== env.PROXY_SECRET) {
        return new Response("Unauthorized", { status: 401 });
    }

    const header = "Date,Email,Plan,Domain,Stripe_Customer,Subscription,Amount_USD\n";
    let rows = [header];

    // List all keys with 'sale:' prefix
    const list = await env.LICENSES.list({ prefix: "sale:" });
    for (const key of list.keys) {
        const val = await env.LICENSES.get(key.name);
        if (val) rows.push(val + "\n");
    }

    return new Response(rows.join(""), {
        headers: {
            "Content-Type": "text/csv",
            "Content-Disposition": "attachment; filename=jira-sync-sales.csv",
            "Access-Control-Allow-Origin": "*"
        }
    });
}

/**
 * Handles License verification using Cloudflare KV
 * Checks both individual email and domain-wide licenses
 */
async function handleLicenseCheck(request, env) {
    try {
        const { email } = await request.json();
        console.log(`[License Check] Checking email: ${email}`);
        if (!email) return new Response("Missing email", { status: 400 });

        const emailLower = email.toLowerCase();
        const domain = emailLower.split('@')[1];

        if (!env.LICENSES) {
            // Development/Fallback logic
            return new Response(JSON.stringify({
                allowed: true,
                plan: "pro (KV not bound)",
                status: "active",
                licenseType: "individual",
                features: ["unlimited_syncs"]
            }), {
                headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
            });
        }

        // 1. Check for individual license
        const userData = await env.LICENSES.get(`user:${emailLower}`);
        if (userData) {
            return new Response(userData, {
                headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
            });
        }

        // 2. Check for domain-wide license
        if (domain) {
            const domainData = await env.LICENSES.get(`domain:${domain}`);
            if (domainData) {
                // If domain license exists, the user is allowed!
                const parsedDomain = JSON.parse(domainData);
                return new Response(JSON.stringify({
                    ...parsedDomain,
                    allowed: parsedDomain.status === 'active' || parsedDomain.status === 'trialing',
                    licenseType: "domain",
                    domain: domain
                }), {
                    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
                });
            }
        }

        // 3. Fallback to free plan
        return new Response(JSON.stringify({
            allowed: false,
            plan: "free",
            status: "none",
            email: emailLower,
            message: "No active license found for this user or domain."
        }), {
            headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" }
        });
    } catch (e) {
        return new Response("License API Error: " + e.message, { status: 500 });
    }
}

/**
 * Handles Stripe Session Creation
 */
async function handleStripeSession(request, env) {
    try {
        const body = await request.json();
        const { email, plan, priceId, domain } = body;

        // Real Stripe API implementation using Fetch
        const stripeSecret = env.STRIPE_SECRET_KEY;
        if (!stripeSecret) {
            return new Response(JSON.stringify({ error: "Stripe Secret Key not configured in Worker." }), {
                status: 500,
                headers: { "Access-Control-Allow-Origin": "*" }
            });
        }

        // Create a Stripe Checkout Session
        const params = new URLSearchParams();
        params.append("customer_email", email);
        params.append("payment_method_types[]", "card");
        params.append("mode", "subscription");
        params.append("line_items[0][price]", priceId || "price_1P2b3c4d5e6f");
        params.append("line_items[0][quantity]", "1");

        // Success redirect back to the worker's own landing page
        const baseUrl = new URL(request.url).origin;
        params.append("success_url", `${baseUrl}/success?session_id={CHECKOUT_SESSION_ID}`);
        params.append("cancel_url", `${baseUrl}/cancel`);

        params.append("metadata[email]", email);
        params.append("metadata[plan]", plan || "pro");
        if (domain) {
            params.append("metadata[domain]", domain);
        }
        params.append("subscription_data[metadata][email]", email);
        params.append("subscription_data[metadata][plan]", plan || "pro");
        if (domain) {
            params.append("subscription_data[metadata][domain]", domain);
        }

        const stripeResponse = await fetch("https://api.stripe.com/v1/checkout/sessions", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${stripeSecret}`,
                "Content-Type": "application/x-www-form-urlencoded"
            },
            body: params.toString()
        });

        const session = await stripeResponse.json();

        if (!stripeResponse.ok) {
            throw new Error(session.error ? session.error.message : "Stripe API Error");
        }

        return new Response(JSON.stringify({
            message: "Stripe session created.",
            url: session.url,
            sessionId: session.id,
            publishableKey: env.STRIPE_PUBLISHABLE_KEY || "pk_0JkeFQXipCRQx1VBmDgCPlHwPCsvC",
            metadata: { email, plan }
        }), {
            headers: {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Headers": "Content-Type"
            }
        });
    } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), {
            status: 500,
            headers: { "Access-Control-Allow-Origin": "*" }
        });
    }
}

/**
 * Handles Stripe Webhooks for Provisioning
 */
async function handleStripeWebhook(request, env) {
    console.log(`[Stripe Webhook] HIT - Method: ${request.method} - URL: ${request.url}`);
    try {
        const signature = request.headers.get("stripe-signature");
        console.log(`[Stripe Webhook] Signature present: ${!!signature}`);
        const bodyText = await request.text();
        const stripeSecret = env.STRIPE_SECRET_KEY;
        const webhookSecret = env.STRIPE_WEBHOOK_SECRET;

        // Note: Full signature verification requires a library or manual crypto
        // For now, we rely on the webhookSecret being present to trust the request
        const isMock = request.headers.get("X-Mock-Request") === "true";
        if (webhookSecret && !signature && !isMock) {
            console.log("[Stripe Webhook] Rejected: Missing signature and not a mock.");
            return new Response("Missing signature", {
                status: 400,
                headers: { "Access-Control-Allow-Origin": "*" }
            });
        }

        const event = JSON.parse(bodyText);
        console.log(`[Stripe Webhook] Received event: ${event.type}`);
        console.log(`[Stripe Webhook] Body Preview: ${bodyText.substring(0, 200)}...`);

        if (event.type === "checkout.session.completed") {
            const session = event.data.object;
            const email = session.metadata.email || session.customer_details.email;
            const plan = session.metadata.plan || "pro";
            const domain = session.metadata.domain;

            if (email && env.LICENSES) {
                const licenseData = {
                    allowed: true,
                    plan: plan,
                    status: "active",
                    domain: domain,
                    expires: null,
                    customer: session.customer,
                    subscription: session.subscription,
                    lastUpdated: new Date().toISOString()
                };

                // 1. Activate individual license
                await env.LICENSES.put(`user:${email.toLowerCase()}`, JSON.stringify(licenseData));

                // 2. Activate domain license if applicable
                if (domain) {
                    await env.LICENSES.put(`domain:${domain.toLowerCase()}`, JSON.stringify({
                        ...licenseData,
                        email: email // Keep track of who bought it
                    }));
                    console.log(`[Provisioning] Domain license activated for ${domain}`);
                }

                console.log(`[Provisioning] License activated for ${email}`);

                // Record Sale for CSV Export
                const saleRecord = [
                    new Date().toISOString(),
                    email,
                    plan,
                    domain || "Personal",
                    session.customer,
                    session.subscription,
                    (session.amount_total / 100).toFixed(2)
                ].join(",");

                // Store with a unique key for easy listing
                const saleKey = `sale:${Date.now()}:${email}`;
                await env.LICENSES.put(saleKey, saleRecord);
            }
        }

        if (event.type === "customer.subscription.deleted") {
            const subscription = event.data.object;
            const email = subscription.metadata.email;

            if (email && env.LICENSES) {
                // We could delete or just mark as inactive
                const existing = await env.LICENSES.get(`user:${email.toLowerCase()}`);
                if (existing) {
                    const data = JSON.parse(existing);
                    data.allowed = false;
                    data.status = "canceled";
                    await env.LICENSES.put(`user:${email.toLowerCase()}`, JSON.stringify(data));
                    console.log(`[De-provisioning] License deactivated for ${email}`);
                }
            }
        }

        return new Response(JSON.stringify({ received: true }), {
            headers: {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*"
            }
        });
    } catch (e) {
        console.error(`[Stripe Webhook Error] ${e.message}`);
        return new Response("Webhook Error: " + e.message, {
            status: 500,
            headers: { "Access-Control-Allow-Origin": "*" }
        });
    }
}

/**
 * Handles Jira API Proxying
 */
async function handleProxy(request, env) {
    const proxySecret = env.PROXY_SECRET || "your-shared-secret-here";
    const url = new URL(request.url);
    const targetUrl = url.searchParams.get("target");

    const incomingSecret = request.headers.get("X-Proxy-Secret");
    if (incomingSecret !== proxySecret) {
        return new Response("Unauthorized.", { status: 401 });
    }

    try {
        const target = new URL(targetUrl);
        const hostname = target.hostname;

        if (!hostname.endsWith(ALLOWED_DOMAIN_SUFFIX)) {
            return new Response("Forbidden. Target domain not allowed.", { status: 403 });
        }

        // --- Circuit Breaker: Daily Limit Per Domain ---
        if (env.LICENSES) {
            const today = new Date().toISOString().split('T')[0];
            const statsKey = `domain_stats:${hostname}:${today}`;

            // Get current count
            let currentCalls = parseInt(await env.LICENSES.get(statsKey) || "0");

            if (currentCalls >= 100000) {
                console.error(`[Circuit Breaker] Limit hit for ${hostname}: ${currentCalls} calls today.`);
                return new Response(JSON.stringify({
                    error: "Daily API limit reached for this domain.",
                    limit: 100000,
                    domain: hostname
                }), {
                    status: 429,
                    headers: {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    }
                });
            }

            // Increment count
            await env.LICENSES.put(statsKey, (currentCalls + 1).toString(), { expirationTtl: 172800 }); // 48h expiry
        }
        // ----------------------------------------------

        const proxyRequest = new Request(targetUrl, {
            method: request.method,
            headers: request.headers,
            body: request.body,
            redirect: "follow",
        });

        const response = await fetch(proxyRequest);
        const responseHeaders = new Headers(response.headers);
        responseHeaders.set("Access-Control-Allow-Origin", "*");

        return new Response(response.body, {
            status: response.status,
            statusText: response.statusText,
            headers: responseHeaders,
        });
    } catch (e) {
        return new Response("Proxy Error: " + e.message, {
            status: 500,
            headers: { "Access-Control-Allow-Origin": "*" }
        });
    }
}

/**
 * Serves Static Assets from KV or returns instructions
 */
async function handleStaticAssets(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    // Handle Success/Cancel Landing Pages
    if (path === "/success") {
        return new Response(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Success - Jira Sync Pro</title>
                <style>
                    body { font-family: -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f8fafc; color: #1e293b; }
                    .card { background: white; padding: 2.5rem; border-radius: 1rem; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); text-align: center; max-width: 400px; }
                    .icon { font-size: 3rem; color: #10b981; margin-bottom: 1rem; }
                    h1 { margin: 0; color: #0f172a; }
                    p { color: #64748b; line-height: 1.5; margin: 1rem 0; }
                    .btn { display: inline-block; background: #2563eb; color: white; padding: 0.75rem 1.5rem; border-radius: 0.5rem; text-decoration: none; font-weight: 600; margin-top: 1rem; }
                </style>
            </head>
            <body>
                <div class="card">
                    <div class="icon">✅</div>
                    <h1>Upgrade Successful!</h1>
                    <p>Your Pro license is now active. You can close this tab and return to your Google Sheet.</p>
                    <a href="javascript:window.close()" class="btn">Close Tab</a>
                </div>
            </body>
            </html>
        `, { headers: { "Content-Type": "text/html" } });
    }

    if (path === "/cancel") {
        return new Response(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Canceled - Jira Sync Pro</title>
                <style>
                    body { font-family: -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f8fafc; color: #1e293b; }
                    .card { background: white; padding: 2.5rem; border-radius: 1rem; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); text-align: center; max-width: 400px; }
                    .icon { font-size: 3rem; color: #64748b; margin-bottom: 1rem; }
                    h1 { margin: 0; color: #0f172a; }
                    p { color: #64748b; line-height: 1.5; margin: 1rem 0; }
                    .btn { display: inline-block; background: #e2e8f0; color: #1e293b; padding: 0.75rem 1.5rem; border-radius: 0.5rem; text-decoration: none; font-weight: 600; margin-top: 1rem; }
                </style>
            </head>
            <body>
                <div class="card">
                    <div class="icon">🛒</div>
                    <h1>Payment Canceled</h1>
                    <p>No charges were made. You can close this tab and return to the add-on to try again.</p>
                    <a href="javascript:window.close()" class="btn">Close Tab</a>
                </div>
            </body>
            </html>
        `, { headers: { "Content-Type": "text/html" } });
    }

    if (!env.ASSETS) {
        return new Response(`
            <h1>Jira Sync Worker Backend</h1>
            <p>API endpoints available at: /api/license/check, /api/stripe/create-session</p>
        `, { headers: { "Content-Type": "text/html" } });
    }

    const key = path === "/" ? "index.html" : path.slice(1);
    const asset = await env.ASSETS.get(key, { type: "stream" });

    if (asset) {
        let contentType = "text/plain";
        if (key.endsWith(".html")) contentType = "text/html";
        else if (key.endsWith(".css")) contentType = "text/css";
        else if (key.endsWith(".js")) contentType = "text/javascript";
        else if (key.endsWith(".png")) contentType = "image/png";
        else if (key.endsWith(".jpg") || key.endsWith(".jpeg")) contentType = "image/jpeg";

        return new Response(asset, { headers: { "Content-Type": contentType } });
    }

    return new Response("Asset not found.", { status: 404 });
}
