import { readFile, writeFile } from 'node:fs/promises';

const sourcePath = process.argv[2];
const outputPath = process.argv[3];
if (!sourcePath || !outputPath) throw new Error('Usage: node fetch-enterprise-hr-prospects.mjs <captured-curl> <output-json>');

const captured = await readFile(sourcePath, 'utf8');
const session = captured.match(/-b '([^']+)'/)?.[1];
const deviceId = captured.match(/-H 'x-device-id: ([^']+)'/)?.[1];
if (!session || !deviceId) throw new Error('The captured request is missing its session or device ID.');

const campaignId = 151855;
const pageSize = 100;
const pauseMs = 1_500;
const prospects = [];
let expected = null;

for (let offset = 0; expected === null || offset < expected; offset += pageSize) {
  const url = new URL(`https://explee.com/api/app-auto-gtm/segments/${campaignId}/leads`);
  url.search = new URLSearchParams({
    people_offset: String(offset),
    people_limit: String(pageSize),
    companies_offset: '0',
    companies_limit: '1',
  });
  const response = await fetch(url, {
    headers: {
      accept: 'application/json',
      cookie: session,
      referer: `https://explee.com/app-auto-gtm/p/34448/segments/${campaignId}?tab=leads`,
      'x-device-id': deviceId,
    },
  });
  if ([401, 403, 429].includes(response.status)) {
    throw new Error(`Stopped immediately after Explee returned ${response.status} at offset ${offset}.`);
  }
  if (!response.ok) throw new Error(`Explee returned ${response.status} at offset ${offset}.`);
  const page = await response.json();
  expected ??= Number(page.meta?.people_count || 0);
  if (!Array.isArray(page.people)) throw new Error(`Unexpected response at offset ${offset}.`);
  prospects.push(...page.people.map((person) => ({
    campaign_id: campaignId,
    person_id: String(person.person_id ?? person.id),
    name: person['Full Name'] || null,
    email: null,
    phone: null,
    job_title: person['Job Title'] || null,
    company_name: person.Company || null,
    company_domain: person.company_domain || null,
    linkedin_url: person.linkedin_url || null,
    country: person.Geo || person.company_geo || null,
    status: 'Prospect',
  })));
  console.log(`Fetched ${prospects.length} of ${expected}`);
  if (prospects.length < expected) await new Promise((resolve) => setTimeout(resolve, pauseMs));
}

const unique = [...new Map(prospects.map((prospect) => [prospect.person_id, prospect])).values()];
if (unique.length !== expected) throw new Error(`Expected ${expected} unique prospects but received ${unique.length}.`);
await writeFile(outputPath, JSON.stringify({ campaign_id: campaignId, prospects: unique }));
console.log(`Saved ${unique.length} prospects.`);
