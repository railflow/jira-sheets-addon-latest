# Firebase License Management API

This Firebase project handles license verification and Stripe subscription management for the Jira Sync Google Sheets Add-on.

## Architecture

```
┌─────────────────────┐     ┌──────────────────────┐     ┌─────────────────┐
│ Google Sheets Addon │────▶│  Firebase Functions  │────▶│   Firestore     │
│                     │     │                      │     │   (licenses)    │
└─────────────────────┘     └──────────────────────┘     └─────────────────┘
                                      ▲
                                      │
                            ┌─────────┴─────────┐
                            │  Stripe Webhooks  │
                            │  (auto-provision) │
                            └───────────────────┘
```

## Setup

### 1. Prerequisites

- Node.js 18+
- Firebase CLI: `npm install -g firebase-tools`
- Stripe account with API keys

### 2. Firebase Setup

```bash
# Login to Firebase
firebase login

# Initialize project (or use existing)
firebase projects:create jira-sync-license

# Set project
firebase use jira-sync-license
```

### 3. Configure Stripe Keys

```bash
# Set Stripe secret key
firebase functions:config:set stripe.secret_key="sk_live_xxxxx"

# Set webhook secret (get this after creating webhook in Stripe)
firebase functions:config:set stripe.webhook_secret="whsec_xxxxx"
```

### 4. Deploy

```bash
cd functions
npm install
cd ..
firebase deploy
```

## API Endpoints

### Check License
```bash
GET/POST https://us-central1-YOUR_PROJECT.cloudfunctions.net/checkLicense
```

**Request:**
```json
{ "email": "user@example.com" }
```

**Response:**
```json
{
  "allowed": true,
  "plan": "pro",
  "status": "active",
  "features": ["unlimited_syncs", "priority_support"],
  "expiresAt": "2025-03-01T00:00:00.000Z"
}
```

### Create Checkout Session
```bash
POST https://us-central1-YOUR_PROJECT.cloudfunctions.net/createCheckoutSession
```

**Request:**
```json
{
  "email": "user@example.com",
  "priceId": "price_xxxxx",
  "successUrl": "https://your-site.com/success",
  "cancelUrl": "https://your-site.com/cancel"
}
```

### Create Customer Portal Session
```bash
POST https://us-central1-YOUR_PROJECT.cloudfunctions.net/createPortalSession
```

**Request:**
```json
{
  "email": "user@example.com",
  "returnUrl": "https://your-site.com/account"
}
```

## Stripe Configuration

### 1. Create Product & Prices

In [Stripe Dashboard](https://dashboard.stripe.com/products):

1. Create a Product: "Jira Sync Pro"
2. Add Prices:
   - Monthly: $9.99/month
   - Yearly: $99/year (save ~17%)

### 2. Configure Webhook

In [Stripe Dashboard → Webhooks](https://dashboard.stripe.com/webhooks):

1. Add endpoint: `https://us-central1-YOUR_PROJECT.cloudfunctions.net/stripeWebhook`
2. Select events:
   - `checkout.session.completed`
   - `customer.subscription.created`
   - `customer.subscription.updated`
   - `customer.subscription.deleted`
   - `invoice.paid`
   - `invoice.payment_failed`
3. Copy the webhook signing secret

### 3. Customer Portal

Enable at [Stripe Dashboard → Settings → Billing → Customer Portal](https://dashboard.stripe.com/settings/billing/portal):

- Allow customers to update payment methods
- Allow customers to cancel subscriptions
- Configure branding

## Firestore Structure

```
licenses/
  └── user@example.com/
        ├── email: "user@example.com"
        ├── stripeCustomerId: "cus_xxxxx"
        ├── stripeSubscriptionId: "sub_xxxxx"
        ├── status: "active" | "trialing" | "past_due" | "canceled" | "expired"
        ├── plan: "pro" | "enterprise"
        ├── features: ["unlimited_syncs", "priority_support", ...]
        ├── currentPeriodStart: Timestamp
        ├── currentPeriodEnd: Timestamp
        ├── cancelAtPeriodEnd: boolean
        ├── createdAt: Timestamp
        └── updatedAt: Timestamp
```

## Local Development

```bash
# Start emulators
firebase emulators:start

# Test webhook locally (use Stripe CLI)
stripe listen --forward-to localhost:5001/YOUR_PROJECT/us-central1/stripeWebhook
```

## Manual License Management

To manually add/update a license:

```javascript
// In Firebase console or via Admin SDK
db.collection('licenses').doc('user@example.com').set({
  email: 'user@example.com',
  status: 'active',
  plan: 'pro',
  features: ['unlimited_syncs', 'priority_support'],
  createdAt: new Date(),
  updatedAt: new Date()
});
```

## Troubleshooting

### Webhook Not Receiving Events
- Verify webhook URL is correct
- Check Stripe webhook logs for errors
- Ensure Cloud Functions have proper permissions

### License Check Returns False
- Verify email is exactly as stored (case-insensitive)
- Check subscription status in Stripe Dashboard
- View Firestore document for the email

### CORS Errors
- The functions include CORS handling for all origins
- For production, restrict to your domain
