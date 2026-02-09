let lastAlertAtMs = 0;

function parsePositiveInt(value, fallback) {
  const parsed = Number.parseInt(value || '', 10);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return fallback;
  }
  return parsed;
}

function getAlertConfig() {
  return {
    slackWebhookUrl: (process.env.ALERT_SLACK_WEBHOOK_URL || '').trim(),
    sendgridApiKey: (process.env.SENDGRID_API_KEY || '').trim(),
    alertEmailTo: (process.env.ALERT_EMAIL_TO || '').trim(),
    alertEmailFrom: (process.env.ALERT_EMAIL_FROM || '').trim(),
    cooldownSeconds: parsePositiveInt(process.env.ALERT_COOLDOWN_SECONDS, 600),
  };
}

function canSendAlert(cooldownSeconds) {
  const now = Date.now();
  if (now - lastAlertAtMs < cooldownSeconds * 1000) {
    return false;
  }
  lastAlertAtMs = now;
  return true;
}

function buildAlertText(title, details) {
  const detailText = typeof details === 'string' ? details : JSON.stringify(details);
  return `${title}\n${detailText}`;
}

async function sendSlackAlert(webhookUrl, text) {
  const response = await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ text }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Slack alert failed ${response.status}: ${body}`);
  }
}

async function sendSendgridEmail(apiKey, to, from, subject, text) {
  const response = await fetch('https://api.sendgrid.com/v3/mail/send', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      personalizations: [{ to: [{ email: to }] }],
      from: { email: from },
      subject,
      content: [{ type: 'text/plain', value: text }],
    }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`SendGrid alert failed ${response.status}: ${body}`);
  }
}

async function notifyFailure(title, details) {
  const cfg = getAlertConfig();
  const hasSlack = Boolean(cfg.slackWebhookUrl);
  const hasEmail = Boolean(cfg.sendgridApiKey && cfg.alertEmailTo && cfg.alertEmailFrom);

  if (!hasSlack && !hasEmail) {
    return;
  }

  if (!canSendAlert(cfg.cooldownSeconds)) {
    return;
  }

  const text = buildAlertText(title, details);
  const tasks = [];

  if (hasSlack) {
    tasks.push(
      sendSlackAlert(cfg.slackWebhookUrl, text).catch((error) => {
        console.error('alerts.slack.error', { message: error.message });
      }),
    );
  }

  if (hasEmail) {
    tasks.push(
      sendSendgridEmail(
        cfg.sendgridApiKey,
        cfg.alertEmailTo,
        cfg.alertEmailFrom,
        title,
        text,
      ).catch((error) => {
        console.error('alerts.email.error', { message: error.message });
      }),
    );
  }

  await Promise.all(tasks);
}

module.exports = {
  notifyFailure,
};
