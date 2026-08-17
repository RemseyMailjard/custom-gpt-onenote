const { app } = require('@azure/functions');
const Stripe = require('stripe');
const sgMail = require('@sendgrid/mail');
const crypto = require('crypto');
const { createLicense } = require('../lib/licenses');

function buildWelcomeEmail({ to, licenseCode }) {
  const proxyBase = process.env.PROXY_BASE_URL || 'https://remsey-onenote-proxy-z3oh.azurewebsites.net';
  const serverUrl = `${proxyBase}/lic/${licenseCode}`;

  return {
    to,
    from: process.env.SENDGRID_FROM_EMAIL,
    subject: 'Your OneNote GPT license is ready',
    text: `Thanks for subscribing to OneNote Assistant.

Your personal server URL (paste this into the "Servers" line of the OpenAPI schema when you build your GPT):

${serverUrl}

Follow the setup guide here to build your GPT with this URL:
${process.env.SETUP_GUIDE_URL || 'https://your-site.example/setup-guide/'}

Keep this URL private — it's tied to your subscription.`,
    html: `<p>Thanks for subscribing to OneNote Assistant.</p>
<p>Your personal server URL (paste this into the <strong>Servers</strong> line of the OpenAPI schema when you build your GPT):</p>
<p><code>${serverUrl}</code></p>
<p>Follow the setup guide here to build your GPT with this URL:<br>
<a href="${process.env.SETUP_GUIDE_URL || 'https://your-site.example/setup-guide/'}">${process.env.SETUP_GUIDE_URL || 'https://your-site.example/setup-guide/'}</a></p>
<p>Keep this URL private — it's tied to your subscription.</p>`
  };
}

app.http('stripeWebhook', {
  route: 'stripe/webhook',
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);
    const signature = request.headers.get('stripe-signature');
    const rawBody = await request.text();

    let event;
    try {
      event = stripe.webhooks.constructEvent(rawBody, signature, process.env.STRIPE_WEBHOOK_SECRET);
    } catch (err) {
      context.error('Stripe signature verification failed', err.message);
      return { status: 400, jsonBody: { error: { message: 'Invalid signature' } } };
    }

    if (event.type !== 'checkout.session.completed') {
      return { status: 200, jsonBody: { received: true } };
    }

    const session = event.data.object;
    const email = session.customer_details?.email || session.customer_email;
    if (!email) {
      context.error('Checkout session has no customer email', session.id);
      return { status: 400, jsonBody: { error: { message: 'Missing customer email' } } };
    }

    const licenseCode = crypto.randomBytes(16).toString('hex');

    await createLicense({
      code: licenseCode,
      email,
      stripeCustomerId: session.customer,
      stripeSessionId: session.id
    });

    if (process.env.SENDGRID_API_KEY) {
      sgMail.setApiKey(process.env.SENDGRID_API_KEY);
      try {
        await sgMail.send(buildWelcomeEmail({ to: email, licenseCode }));
      } catch (err) {
        context.error('Failed to send license email', err.message);
      }
    } else {
      context.warn('SENDGRID_API_KEY not set — license created but no email sent', { email, licenseCode });
    }

    return { status: 200, jsonBody: { received: true } };
  }
});
