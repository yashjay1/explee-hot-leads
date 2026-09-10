import { readFile } from 'node:fs/promises';

const sourcePath = process.argv[2];
const apiBase = process.argv[3];
const token = process.env.SIGNAL_DESK_ID_TOKEN;
if (!sourcePath || !apiBase || !token) throw new Error('Source path, API URL, and SIGNAL_DESK_ID_TOKEN are required.');

const payload = JSON.parse(await readFile(sourcePath, 'utf8'));
const prospects = payload.prospects;
if (!Array.isArray(prospects) || payload.campaign_id !== 151855) throw new Error('Invalid Enterprise HR export.');

for (let index = 0; index < prospects.length; index += 250) {
  const batch = prospects.slice(index, index + 250);
  const response = await fetch(`${apiBase.replace(/\/$/, '')}/api/prospects/import`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ prospects: batch }),
  });
  if (!response.ok) throw new Error(`Import stopped at ${index}: ${response.status} ${await response.text()}`);
  console.log(`Imported ${Math.min(index + batch.length, prospects.length)} of ${prospects.length}`);
}
