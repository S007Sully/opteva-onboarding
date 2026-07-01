export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const RESEND_API_KEY = process.env.RESEND_API_KEY;
  if (!RESEND_API_KEY) return res.status(500).json({ error: 'RESEND_API_KEY not configured.' });

  const d = req.body;

  const socials = [
    d.instagram && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:120px;">Instagram</td><td style="padding:6px 0;font-size:13px;font-weight:600;">@${d.instagram}</td></tr>`,
    d.facebook  && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Facebook</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.facebook}</td></tr>`,
    d.tiktok    && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">TikTok</td><td style="padding:6px 0;font-size:13px;font-weight:600;">@${d.tiktok}</td></tr>`,
    d.linkedin  && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">LinkedIn</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.linkedin}</td></tr>`,
    d.youtube   && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">YouTube</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.youtube}</td></tr>`,
    d.xtwitter  && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">X / Twitter</td><td style="padding:6px 0;font-size:13px;font-weight:600;">@${d.xtwitter}</td></tr>`,
    d.pinterest && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Pinterest</td><td style="padding:6px 0;font-size:13px;font-weight:600;">@${d.pinterest}</td></tr>`,
    d.gmb       && `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Google Business</td><td style="padding:6px 0;font-size:13px;font-weight:600;"><a href="${d.gmb}" style="color:#0D9488;">${d.gmb}</a></td></tr>`,
  ].filter(Boolean).join('');

  const webUrls = (d.webUrls || []).map(u =>
    `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Link</td><td style="padding:6px 0;font-size:13px;"><a href="${u}" style="color:#0D9488;">${u}</a></td></tr>`
  ).join('');

  const html = `
<!DOCTYPE html><html><body style="font-family:Inter,Arial,sans-serif;background:#F8FAFC;margin:0;padding:24px;">
<div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #E2E8F0;">
  <div style="background:#0F172A;padding:24px 28px;display:flex;align-items:center;gap:12px;">
    <div style="font-size:20px;font-weight:800;color:#fff;letter-spacing:-0.02em;">Opteva</div>
    <div style="font-size:13px;color:rgba(255,255,255,0.5);margin-left:4px;">New Client Onboarding</div>
  </div>
  <div style="padding:28px;">
    <div style="background:#F0FDFA;border:1px solid #99F6E4;border-radius:10px;padding:16px 20px;margin-bottom:24px;">
      <div style="font-size:18px;font-weight:800;color:#0F172A;margin-bottom:2px;">${d.firstName} ${d.lastName}</div>
      <div style="font-size:14px;color:#0D9488;font-weight:600;">${d.bizName}</div>
    </div>

    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
      <tr style="background:#F8FAFC;"><td colspan="2" style="padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;border-radius:6px;">Contact</td></tr>
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:120px;">Email</td><td style="padding:6px 0;font-size:13px;font-weight:600;"><a href="mailto:${d.email}" style="color:#0D9488;">${d.email}</a></td></tr>
      ${d.phone ? `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Phone</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.phone}</td></tr>` : ''}
      ${d.website ? `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Website</td><td style="padding:6px 0;font-size:13px;font-weight:600;"><a href="${d.website}" style="color:#0D9488;">${d.website}</a></td></tr>` : ''}
    </table>

    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
      <tr style="background:#F8FAFC;"><td colspan="2" style="padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;border-radius:6px;">Brand</td></tr>
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:120px;">Industry</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.industry}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Audience</td><td style="padding:6px 0;font-size:13px;">${d.audience}</td></tr>
      ${d.brandVoice ? `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Brand Voice</td><td style="padding:6px 0;font-size:13px;">${d.brandVoice}</td></tr>` : ''}
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Colors</td><td style="padding:6px 0;font-size:13px;"><span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:${d.color1};vertical-align:middle;margin-right:5px;border:1px solid #E2E8F0;"></span>${d.color1}&nbsp;&nbsp;<span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:${d.color2};vertical-align:middle;margin-right:5px;border:1px solid #E2E8F0;"></span>${d.color2}</td></tr>
    </table>

    ${socials ? `<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
      <tr style="background:#F8FAFC;"><td colspan="2" style="padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;border-radius:6px;">Socials</td></tr>
      ${socials}
    </table>` : ''}

    ${webUrls ? `<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
      <tr style="background:#F8FAFC;"><td colspan="2" style="padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;border-radius:6px;">Web Presence to Scrape</td></tr>
      ${webUrls}
    </table>` : ''}

    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
      <tr style="background:#F8FAFC;"><td colspan="2" style="padding:8px 12px;font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;border-radius:6px;">Goals</td></tr>
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;width:120px;">Goals</td><td style="padding:6px 0;font-size:13px;">${d.goals}</td></tr>
      <tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Frequency</td><td style="padding:6px 0;font-size:13px;font-weight:600;">${d.frequency}</td></tr>
      ${d.notes ? `<tr><td style="padding:6px 0;color:#64748B;font-size:13px;">Notes</td><td style="padding:6px 0;font-size:13px;">${d.notes}</td></tr>` : ''}
    </table>

    ${d.logo ? `<div style="margin-bottom:24px;"><div style="font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94A3B8;margin-bottom:10px;background:#F8FAFC;padding:8px 12px;border-radius:6px;">Logo</div><img src="${d.logo}" style="max-width:180px;max-height:80px;object-fit:contain;border:1px solid #E2E8F0;border-radius:8px;padding:8px;"></div>` : ''}

    <div style="background:#0F172A;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:13px;color:rgba(255,255,255,0.6);margin-bottom:4px;">Submitted via</div>
      <div style="font-size:14px;font-weight:700;color:#0D9488;">onboarding.opteva.ai</div>
    </div>
  </div>
</div>
</body></html>`;

  const emailPayload = {
    from: 'Opteva Onboarding <onboarding@opteva.ai>',
    to: ['hello@opteva.ai'],
    reply_to: d.email,
    subject: `New Client: ${d.firstName} ${d.lastName} — ${d.bizName}`,
    html,
  };

  try {
    const r = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${RESEND_API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(emailPayload),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.message || 'Resend error');
    return res.status(200).json({ success: true });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
