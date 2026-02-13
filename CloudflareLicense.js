/**
 * CloudflareLicense.js
 * Handles interactions with Cloudflare Worker for license management.
 */

/**
 * Gets the Cloudflare Worker URL from script properties
 */
function getWorkerUrl() {
    return PropertiesService.getScriptProperties().getProperty('CLOUDFLARE_WORKER_URL') || CLOUDFLARE_WORKER_URL;
}

/**
 * Checks license status for a user via Cloudflare Worker
 */
/**
 * Checks license status for a user via Cloudflare Worker.
 * Caches the result for 24 hours as per requirements.
 */
function getLicenseStatus(email, forceRefresh = false) {
    if (!email) return { status: 'none', plan: 'free', allowed: false };

    const props = PropertiesService.getUserProperties();
    const cacheKey = 'LICENSE_CACHE_' + Utilities.base64Encode(email);
    const lastCheckKey = 'LICENSE_LAST_CHECK_' + Utilities.base64Encode(email);

    const cachedData = props.getProperty(cacheKey);
    const lastCheck = props.getProperty(lastCheckKey);
    const now = new Date().getTime();

    // 24 hours in milliseconds
    const ONE_DAY = 24 * 60 * 60 * 1000;

    if (!forceRefresh && cachedData && lastCheck && (now - parseInt(lastCheck)) < ONE_DAY) {
        try {
            return JSON.parse(cachedData);
        } catch (e) {
            console.warn('Failed to parse cached license data');
        }
    }

    const workerUrl = getWorkerUrl();
    if (!workerUrl) {
        console.warn('Cloudflare Worker URL missing');
        return { status: 'none', plan: 'free', allowed: false };
    }

    const url = `${workerUrl}/api/license/check`;

    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify({ email: email }),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            const licenseInfo = {
                ...data,
                allowed: data.allowed || false,
                lastVerified: now
            };

            // Cache the result
            props.setProperty(cacheKey, JSON.stringify(licenseInfo));
            props.setProperty(lastCheckKey, now.toString());

            return licenseInfo;
        }

        console.warn('Cloudflare License API returned error:', response.getResponseCode());
        // Fallback to cached data if API fails
        if (cachedData) return JSON.parse(cachedData);

        return { status: 'none', plan: 'free', allowed: false };
    } catch (e) {
        console.error('Cloudflare License Check Error:', e);
        if (cachedData) return JSON.parse(cachedData);
        return { status: 'none', plan: 'free', allowed: false };
    }
}

/**
 * Adds a license (Admin feature - would need a secret in real app)
 */
function addLicense(email, plan, durationDays = 365) {
    // For now, this is a placeholder as updating KV requires a different API or secret handling
    console.log('Manual license addition for', email, plan);
    return { success: false, message: 'License addition via Apps Script is not yet implemented for Cloudflare.' };
}
