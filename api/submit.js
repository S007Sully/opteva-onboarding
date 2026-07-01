export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const GHL_API_KEY = process.env.GHL_API_KEY;
  const GHL_LOCATION_ID = process.env.GHL_LOCATION_ID;
  if (!GHL_API_KEY || !GHL_LOCATION_ID) {
    return res.status(500).json({ error: 'GHL credentials not configured.' });
  }

  const d = req.body;

  const socialNotes = [
    d.instagram  ? `Instagram: @${d.instagram}`       : null,
    d.facebook   ? `Facebook: ${d.facebook}`           : null,
    d.tiktok     ? `TikTok: @${d.tiktok}`             : null,
    d.linkedin   ? `LinkedIn: ${d.linkedin}`           : null,
    d.youtube    ? `YouTube: ${d.youtube}`             : null,
    d.xtwitter   ? `X/Twitter: @${d.xtwitter}`        : null,
    d.pinterest  ? `Pinterest: @${d.pinterest}`        : null,
    d.gmb        ? `Google Business: ${d.gmb}`         : null,
  ].filter(Boolean).join('\n');

  const webNotes = (d.webUrls || []).length
    ? `\nWeb to scrape:\n${d.webUrls.join('\n')}`
    : '';

  const brandNotes = [
    `Business: ${d.bizName}`,
    `Industry: ${d.industry}`,
    d.audience   ? `Audience: ${d.audience}`           : null,
    d.brandVoice ? `Voice: ${d.brandVoice}`            : null,
    `Colors: ${d.color1} / ${d.color2}`,
    `Goals: ${d.goals}`,
    `Frequency: ${d.frequency}`,
    d.notes      ? `Notes: ${d.notes}`                 : null,
    socialNotes  ? `\nSocials:\n${socialNotes}`        : null,
    webNotes     ? webNotes                            : null,
  ].filter(Boolean).join('\n');

  const contact = {
    locationId: GHL_LOCATION_ID,
    firstName: d.firstName,
    lastName: d.lastName,
    email: d.email,
    phone: d.phone || undefined,
    website: d.website || undefined,
    source: 'Opteva Onboarding Form',
    tags: ['onboarding', 'new-client'],
    customFields: [
      { key: 'business_name', field_value: d.bizName },
      { key: 'industry',      field_value: d.industry },
      { key: 'goals',         field_value: d.goals },
      { key: 'notes',         field_value: brandNotes },
    ].filter(f => f.field_value),
  };

  try {
    const r = await fetch('https://services.leadconnectorhq.com/contacts/', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${GHL_API_KEY}`,
        'Content-Type': 'application/json',
        'Version': '2021-07-28',
      },
      body: JSON.stringify(contact),
    });

    const data = await r.json();

    if (!r.ok) {
      console.error('GHL error:', data);
      return res.status(500).json({ error: data.message || 'GHL API error' });
    }

    return res.status(200).json({ success: true, contactId: data.contact?.id });
  } catch (err) {
    console.error('Submit error:', err);
    return res.status(500).json({ error: err.message });
  }
}
