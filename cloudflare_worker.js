/**
 * Cloudflare Worker for Jira Proxy & License Management
 * Storage: Cloudflare D1 (SQLite) via env.DB binding
 */

const ALLOWED_DOMAIN_SUFFIX = ".atlassian.net";
const CORS = { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" };

export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);
        const path = url.pathname;

        // CORS preflight
        if (request.method === "OPTIONS") {
            return new Response(null, {
                headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
                    "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Atlassian-Token, X-Proxy-Secret",
                },
            });
        }

        if (path === "/api/license/check")        return handleLicenseCheck(request, env, ctx);
        if (path === "/api/stripe/create-session") return handleStripeSession(request, env);
        if (path === "/api/stripe/webhook")        return handleStripeWebhook(request, env);
        if (path === "/api/stripe/portal")         return handleStripePortal(request, env);
        if (path === "/api/admin/sales-csv")       return handleSalesExport(request, env);

        if (path === "/api/debug") {
            return new Response(JSON.stringify({
                hasDB: !!env.DB,
                time: new Date().toISOString(),
            }), { headers: CORS });
        }

        if (url.searchParams.has("target")) return handleProxy(request, env);

        return handleStaticAssets(request, env);
    },
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

function okJson(data)       { return new Response(JSON.stringify(data), { headers: CORS }); }
function errJson(msg, code) { return new Response(JSON.stringify({ error: msg }), { status: code, headers: CORS }); }

// In-memory rate limiter (per worker instance; provides basic protection against burst abuse)
const _rlMap = new Map();
function rateLimitCheck(key, maxRequests, windowMs) {
    const now = Date.now();
    let entry = _rlMap.get(key);
    if (!entry || now > entry.resetAt) {
        entry = { count: 0, resetAt: now + windowMs };
    }
    entry.count++;
    _rlMap.set(key, entry);
    return entry.count <= maxRequests;
}

function getClientIp(request) {
    return request.headers.get("CF-Connecting-IP") ||
           (request.headers.get("X-Forwarded-For") || "").split(",")[0].trim() ||
           "unknown";
}

/**
 * Login event upsert — registered with ctx.waitUntil so the D1 write
 * completes even after the response is sent.
 */
function logLogin(env, ctx, email, plan) {
    if (!env.DB) return;
    const now = new Date().toISOString();
    const p = env.DB.prepare(`
        INSERT INTO login_events (email, first_seen, last_seen, visit_count, plan)
        VALUES (?, ?, ?, 1, ?)
        ON CONFLICT(email) DO UPDATE SET
            last_seen   = excluded.last_seen,
            visit_count = visit_count + 1,
            plan        = excluded.plan
    `).bind(email, now, now, plan || "free").run().catch(() => {});
    ctx?.waitUntil(p);
}

// ─── License Check ────────────────────────────────────────────────────────────

async function handleLicenseCheck(request, env, ctx) {
    // Rate limit: 20 requests per minute per IP
    const ip = getClientIp(request);
    if (!rateLimitCheck(`lc:${ip}`, 20, 60_000)) {
        return new Response(JSON.stringify({ error: "Too many requests." }), { status: 429, headers: { "Content-Type": "application/json" } });
    }

    try {
        const body = await request.json();
        const email = body?.email;
        if (!email) return errJson("Missing email", 400);

        const emailLower = email.toLowerCase();
        const domain = emailLower.split('@')[1];

        if (!env.DB) {
            return okJson({ allowed: true, plan: "pro (DB not bound)", status: "active", licenseType: "individual", features: ["unlimited_syncs"] });
        }

        // 1. Individual license
        const user = await env.DB
            .prepare("SELECT * FROM user_licenses WHERE email = ?")
            .bind(emailLower).first();

        if (user) {
            logLogin(env, ctx, emailLower, user.plan);
            return okJson({
                allowed:      !!user.allowed,
                plan:         user.plan,
                status:       user.status,
                licenseType:  "individual",
                customer:     user.customer_id,
                subscription: user.subscription_id,
                priceId:      user.price_id,
                amount:       user.amount,
                renewsAt:     user.renews_at,
                domain:       user.domain,
                lastUpdated:  user.last_updated,
            });
        }

        // 2. Domain license
        if (domain) {
            const dom = await env.DB
                .prepare("SELECT * FROM domain_licenses WHERE domain = ?")
                .bind(domain).first();

            if (dom) {
                const allowed = dom.status === "active" || dom.status === "trialing";
                logLogin(env, ctx, emailLower, dom.plan);
                return okJson({
                    allowed,
                    plan:         dom.plan,
                    status:       dom.status,
                    licenseType:  "domain",
                    domain,
                    customer:     dom.customer_id,
                    subscription: dom.subscription_id,
                    priceId:      dom.price_id,
                    amount:       dom.amount,
                    renewsAt:     dom.renews_at,
                    seats:        dom.seats,
                });
            }
        }

        // 3. Free / no license
        logLogin(env, ctx, emailLower, "free");
        return okJson({ allowed: false, plan: "free", status: "none", email: emailLower, message: "No active license found." });

    } catch (e) {
        console.error("[License Check Error]", e.message);
        return new Response(JSON.stringify({ error: "Internal server error." }), { status: 500, headers: { "Content-Type": "application/json" } });
    }
}

// ─── Sales Export ─────────────────────────────────────────────────────────────

async function handleSalesExport(request, env) {
    const authHeader = request.headers.get("Authorization") || "";
    const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
    if (!token || !env.PROXY_SECRET || token !== env.PROXY_SECRET) {
        return new Response("Unauthorized", { status: 401 });
    }

    const { results } = await env.DB.prepare("SELECT * FROM sales ORDER BY created_at DESC").all();
    const header = "Date,Email,Plan,Domain,Stripe_Customer,Subscription,Amount_USD\n";
    const rows = (results || []).map(r =>
        [r.created_at, r.email, r.plan, r.domain || "Personal", r.customer_id, r.subscription_id, (r.amount || 0).toFixed(2)].join(",")
    );

    return new Response(header + rows.join("\n"), {
        headers: {
            "Content-Type": "text/csv",
            "Content-Disposition": "attachment; filename=jira-sync-sales.csv",
        },
    });
}

// ─── Stripe Session ───────────────────────────────────────────────────────────

async function handleStripeSession(request, env) {
    try {
        const { email, plan, priceId, domain } = await request.json();
        const stripeSecret = env.STRIPE_SECRET_KEY;
        if (!stripeSecret) return errJson("Stripe Secret Key not configured.", 500);

        const params = new URLSearchParams();
        params.append("customer_email", email);
        params.append("payment_method_types[]", "card");
        params.append("mode", "subscription");
        params.append("line_items[0][price]", priceId || "price_1P2b3c4d5e6f");
        params.append("line_items[0][quantity]", "1");

        const baseUrl = new URL(request.url).origin;
        params.append("success_url", `${baseUrl}/success?session_id={CHECKOUT_SESSION_ID}`);
        params.append("cancel_url", `${baseUrl}/cancel`);
        params.append("metadata[email]", email);
        params.append("metadata[plan]", plan || "pro");
        if (domain) params.append("metadata[domain]", domain);
        params.append("subscription_data[metadata][email]", email);
        params.append("subscription_data[metadata][plan]", plan || "pro");
        if (domain) params.append("subscription_data[metadata][domain]", domain);

        const res = await fetch("https://api.stripe.com/v1/checkout/sessions", {
            method: "POST",
            headers: { "Authorization": `Bearer ${stripeSecret}`, "Content-Type": "application/x-www-form-urlencoded" },
            body: params.toString(),
        });

        const session = await res.json();
        if (!res.ok) throw new Error(session.error?.message || "Stripe API Error");

        return okJson({ message: "Session created.", url: session.url, sessionId: session.id, publishableKey: env.STRIPE_PUBLISHABLE_KEY, metadata: { email, plan } });
    } catch (e) {
        console.error("[Stripe Session Error]", e.message);
        return errJson("Failed to create checkout session.", 500);
    }
}

// ─── Stripe Webhook ───────────────────────────────────────────────────────────

async function handleStripeWebhook(request, env) {
    console.log(`[Webhook] HIT ${request.method} ${request.url}`);
    try {
        const signature   = request.headers.get("stripe-signature");
        const bodyText    = await request.text();

        if (env.STRIPE_WEBHOOK_SECRET && !signature) {
            return new Response("Missing signature", { status: 400, headers: CORS });
        }

        const event = JSON.parse(bodyText);
        console.log(`[Webhook] Event: ${event.type}`);
        const stripeSecret = env.STRIPE_SECRET_KEY;

        // ── checkout.session.completed ──────────────────────────────────────
        if (event.type === "checkout.session.completed") {
            const session = event.data.object;
            const email   = session.metadata?.email || session.customer_details?.email;
            const plan    = session.metadata?.plan || "pro";
            const domain  = session.metadata?.domain || null;
            if (!email) { console.error("[Webhook] No email in session"); return okJson({ received: true }); }

            // Fetch subscription for renewal date + price
            let renewsAt = null, priceId = null, amount = null;
            if (session.subscription && stripeSecret) {
                try {
                    const subRes = await fetch(`https://api.stripe.com/v1/subscriptions/${session.subscription}`, {
                        headers: { "Authorization": `Bearer ${stripeSecret}` },
                    });
                    if (subRes.ok) {
                        const sub = await subRes.json();
                        renewsAt = new Date(sub.current_period_end * 1000).toISOString();
                        priceId  = sub.items?.data?.[0]?.price?.id || null;
                        amount   = sub.items?.data?.[0]?.price?.unit_amount
                            ? sub.items.data[0].price.unit_amount / 100
                            : (session.amount_total / 100);
                    }
                } catch (e) { console.error("[Webhook] Sub fetch failed:", e.message); }
            }

            const now = new Date().toISOString();

            // Upsert individual license
            await env.DB.prepare(`
                INSERT INTO user_licenses (email, domain, plan, status, allowed, customer_id, subscription_id, price_id, amount, renews_at, last_updated)
                VALUES (?, ?, ?, 'active', 1, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(email) DO UPDATE SET
                    domain = excluded.domain, plan = excluded.plan, status = 'active', allowed = 1,
                    customer_id = excluded.customer_id, subscription_id = excluded.subscription_id,
                    price_id = excluded.price_id, amount = excluded.amount,
                    renews_at = excluded.renews_at, last_updated = excluded.last_updated
            `).bind(email.toLowerCase(), domain || null, plan, session.customer, session.subscription, priceId, amount, renewsAt, now).run();

            // Upsert domain license if applicable
            if (domain) {
                await env.DB.prepare(`
                    INSERT INTO domain_licenses (domain, email, plan, status, allowed, customer_id, subscription_id, price_id, amount, renews_at, last_updated)
                    VALUES (?, ?, ?, 'active', 1, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(domain) DO UPDATE SET
                        email = excluded.email, plan = excluded.plan, status = 'active', allowed = 1,
                        customer_id = excluded.customer_id, subscription_id = excluded.subscription_id,
                        price_id = excluded.price_id, amount = excluded.amount,
                        renews_at = excluded.renews_at, last_updated = excluded.last_updated
                `).bind(domain.toLowerCase(), email, plan, session.customer, session.subscription, priceId, amount, renewsAt, now).run();
                console.log(`[Provisioning] Domain license activated: ${domain}`);
            }

            // Record sale
            await env.DB.prepare(`
                INSERT INTO sales (created_at, email, plan, domain, customer_id, subscription_id, amount)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            `).bind(now, email, plan, domain || null, session.customer, session.subscription, amount || (session.amount_total / 100)).run();

            console.log(`[Provisioning] License activated: ${email}`);
        }

        // ── customer.subscription.updated ───────────────────────────────────
        if (event.type === "customer.subscription.updated") {
            const sub    = event.data.object;
            const email  = sub.metadata?.email;
            const domain = sub.metadata?.domain || null;
            if (email && env.DB) {
                const renewsAt = new Date(sub.current_period_end * 1000).toISOString();
                const priceId  = sub.items?.data?.[0]?.price?.id || null;
                const active   = sub.status === "active" || sub.status === "trialing" ? 1 : 0;
                const now      = new Date().toISOString();
                await env.DB.prepare(`
                    UPDATE user_licenses
                    SET status = ?, allowed = ?, renews_at = ?, price_id = COALESCE(?, price_id), last_updated = ?
                    WHERE email = ?
                `).bind(sub.status, active, renewsAt, priceId, now, email.toLowerCase()).run();
                if (domain) {
                    await env.DB.prepare(`
                        UPDATE domain_licenses
                        SET status = ?, allowed = ?, renews_at = ?, price_id = COALESCE(?, price_id), last_updated = ?
                        WHERE domain = ?
                    `).bind(sub.status, active, renewsAt, priceId, now, domain.toLowerCase()).run();
                    console.log(`[Webhook] Domain license updated for ${domain}: ${sub.status}`);
                }
                console.log(`[Webhook] Subscription updated for ${email}: ${sub.status}`);
            }
        }

        // ── customer.subscription.deleted ───────────────────────────────────
        if (event.type === "customer.subscription.deleted") {
            const sub    = event.data.object;
            const email  = sub.metadata?.email;
            const domain = sub.metadata?.domain || null;
            if (email && env.DB) {
                const now = new Date().toISOString();
                await env.DB.prepare(`
                    UPDATE user_licenses SET allowed = 0, status = 'canceled', last_updated = ? WHERE email = ?
                `).bind(now, email.toLowerCase()).run();
                if (domain) {
                    await env.DB.prepare(`
                        UPDATE domain_licenses SET allowed = 0, status = 'canceled', last_updated = ? WHERE domain = ?
                    `).bind(now, domain.toLowerCase()).run();
                    console.log(`[De-provisioning] Domain license canceled: ${domain}`);
                }
                console.log(`[De-provisioning] License canceled: ${email}`);
            }
        }

        return okJson({ received: true });
    } catch (e) {
        console.error(`[Webhook Error] ${e.message}`);
        return new Response(JSON.stringify({ error: "Webhook processing failed." }), { status: 500, headers: CORS });
    }
}

// ─── Stripe Portal ────────────────────────────────────────────────────────────

async function handleStripePortal(request, env) {
    try {
        const { email, returnUrl } = await request.json();
        if (!email) return errJson("Missing email", 400);
        if (!env.DB)  return errJson("DB not configured", 500);

        const user = await env.DB
            .prepare("SELECT customer_id FROM user_licenses WHERE email = ?")
            .bind(email.toLowerCase()).first();

        if (!user) return errJson("No active subscription found. Please subscribe first.", 404);

        const customerId = user.customer_id;
        if (!customerId || customerId === "cus_mock") {
            return errJson("You are on a trial or mock license. Subscribe to a Pro plan to manage via Stripe.", 400);
        }

        const stripeSecret = env.STRIPE_SECRET_KEY;
        if (!stripeSecret) return errJson("Stripe connection error.", 500);

        const params = new URLSearchParams();
        params.append("customer", customerId);
        if (returnUrl) params.append("return_url", returnUrl);

        const res = await fetch("https://api.stripe.com/v1/billing_portal/sessions", {
            method: "POST",
            headers: { "Authorization": `Bearer ${stripeSecret}`, "Content-Type": "application/x-www-form-urlencoded" },
            body: params.toString(),
        });

        const session = await res.json();
        if (!res.ok) throw new Error(session.error?.message || "Stripe Portal API Error");

        return okJson({ url: session.url });
    } catch (e) {
        console.error(`[Portal Error] ${e.message}`);
        return errJson("Failed to create portal session.", 500);
    }
}

// ─── Jira Proxy ───────────────────────────────────────────────────────────────

async function handleProxy(request, env) {
    const proxySecret = env.PROXY_SECRET;
    if (!proxySecret) return new Response("Server configuration error.", { status: 500 });
    const url = new URL(request.url);
    const targetUrl = url.searchParams.get("target");

    if (request.headers.get("X-Proxy-Secret") !== proxySecret) {
        return new Response("Unauthorized.", { status: 401 });
    }

    try {
        const target   = new URL(targetUrl);
        const hostname = target.hostname;

        if (!hostname.endsWith(ALLOWED_DOMAIN_SUFFIX)) {
            return new Response("Forbidden. Target domain not allowed.", { status: 403 });
        }

        // Circuit breaker — daily limit per Jira domain
        if (env.DB) {
            const today = new Date().toISOString().split("T")[0];

            const row = await env.DB
                .prepare("SELECT call_count FROM domain_stats WHERE hostname = ? AND date = ?")
                .bind(hostname, today).first();

            const currentCalls = row?.call_count || 0;

            if (currentCalls >= 100000) {
                console.error(`[Circuit Breaker] Limit hit for ${hostname}: ${currentCalls} calls today.`);
                return new Response(JSON.stringify({ error: "Daily API limit reached.", limit: 100000, domain: hostname }), {
                    status: 429, headers: CORS,
                });
            }

            await env.DB.prepare(`
                INSERT INTO domain_stats (hostname, date, call_count) VALUES (?, ?, 1)
                ON CONFLICT(hostname, date) DO UPDATE SET call_count = call_count + 1
            `).bind(hostname, today).run();
        }

        const response = await fetch(new Request(targetUrl, {
            method: request.method,
            headers: request.headers,
            body: request.body,
            redirect: "follow",
        }));

        const responseHeaders = new Headers(response.headers);
        responseHeaders.set("Access-Control-Allow-Origin", "*");

        return new Response(response.body, {
            status: response.status,
            statusText: response.statusText,
            headers: responseHeaders,
        });
    } catch (e) {
        console.error("[Proxy Error]", e.message);
        return new Response(JSON.stringify({ error: "Proxy request failed." }), { status: 500, headers: CORS });
    }
}

// ─── Static Assets ────────────────────────────────────────────────────────────

async function handleStaticAssets(request, env) {
    const path = new URL(request.url).pathname;

    if (path === "/success") {
        return new Response(`<!DOCTYPE html><html><head><title>Success - Jira Sync Pro</title>
            <style>body{font-family:-apple-system,sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f8fafc;color:#1e293b}
            .card{background:white;padding:2.5rem;border-radius:1rem;box-shadow:0 10px 25px -5px rgba(0,0,0,.1);text-align:center;max-width:400px}
            .icon{font-size:3rem;color:#10b981;margin-bottom:1rem}h1{margin:0;color:#0f172a}p{color:#64748b;line-height:1.5;margin:1rem 0}
            .btn{display:inline-block;background:#2563eb;color:white;padding:.75rem 1.5rem;border-radius:.5rem;text-decoration:none;font-weight:600;margin-top:1rem}</style>
            </head><body><div class="card"><div class="icon">✅</div><h1>Upgrade Successful!</h1>
            <p>Your license is now active. Close this tab and return to your Google Sheet.</p>
            <a href="javascript:window.close()" class="btn">Close Tab</a></div></body></html>`,
            { headers: { "Content-Type": "text/html" } });
    }

    if (path === "/cancel") {
        return new Response(`<!DOCTYPE html><html><head><title>Canceled - Jira Sync Pro</title>
            <style>body{font-family:-apple-system,sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f8fafc;color:#1e293b}
            .card{background:white;padding:2.5rem;border-radius:1rem;box-shadow:0 10px 25px -5px rgba(0,0,0,.1);text-align:center;max-width:400px}
            .icon{font-size:3rem;color:#64748b;margin-bottom:1rem}h1{margin:0;color:#0f172a}p{color:#64748b;line-height:1.5;margin:1rem 0}
            .btn{display:inline-block;background:#e2e8f0;color:#1e293b;padding:.75rem 1.5rem;border-radius:.5rem;text-decoration:none;font-weight:600;margin-top:1rem}</style>
            </head><body><div class="card"><div class="icon">🛒</div><h1>Payment Canceled</h1>
            <p>No charges were made. Return to the add-on to try again.</p>
            <a href="javascript:window.close()" class="btn">Close Tab</a></div></body></html>`,
            { headers: { "Content-Type": "text/html" } });
    }

    if (!env.ASSETS) {
        return new Response(`<h1>Jira Sync Worker</h1><p>API: /api/license/check, /api/stripe/create-session</p>`, { headers: { "Content-Type": "text/html" } });
    }

    const key   = path === "/" ? "index.html" : path.slice(1);
    const asset = await env.ASSETS.get(key, { type: "stream" });
    if (!asset) return new Response("Asset not found.", { status: 404 });

    const ext = key.split(".").pop();
    const types = { html: "text/html", css: "text/css", js: "text/javascript", png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg" };
    return new Response(asset, { headers: { "Content-Type": types[ext] || "text/plain" } });
}
