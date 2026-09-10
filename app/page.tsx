'use client';
/* oxlint-disable react(react-compiler) -- MSAL synchronizes signed-in state from the browser cache. */

import { useCallback, useEffect, useMemo, useState } from 'react';
import { PublicClientApplication, type AccountInfo } from '@azure/msal-browser';
import { ArrowUpRight, AtSign, Check, ChevronDown, CircleAlert, Clock3, FilePenLine, Flame, Mail, Phone, RefreshCw, Search, Send, Settings2, SlidersHorizontal, Sparkles, UserRound, X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Label } from '@/components/ui/label';

type Lead = {
  person_id: string; campaign_id: number; name: string | null; email: string | null;
  phone: string | null; job_title: string | null; company_name: string | null;
  company_domain: string | null; linkedin_url: string | null; country: string | null;
  why_hot: string | null; became_hot_at: string | null; status?: 'Prospect' | 'New' | 'Drafted' | 'Sent';
};
type DraftSettings = { instruction: string; example: string; signature: string };

const demoLeads: Lead[] = [
  { person_id: 'demo-1', campaign_id: 1812, name: 'Maya Chen', email: 'maya@northstar.example', phone: '+1 (415) 555-0136', job_title: 'VP of Revenue', company_name: 'Northstar Labs', company_domain: 'northstar.example', linkedin_url: 'https://linkedin.com', country: 'US', why_hot: 'This is timely. We are reworking our outbound process this quarter — could you send over a few details and some times next week?', became_hot_at: '2026-09-09T16:24:00Z', status: 'New' },
  { person_id: 'demo-2', campaign_id: 1812, name: 'Jordan Blake', email: 'jordan@hearthside.example', phone: '+44 20 7946 0182', job_title: 'Founder', company_name: 'Hearthside', company_domain: 'hearthside.example', linkedin_url: 'https://linkedin.com', country: 'GB', why_hot: 'Yes, interested. Can you share how pricing works for a team of twelve?', became_hot_at: '2026-09-09T15:02:00Z', status: 'Drafted' },
  { person_id: 'demo-3', campaign_id: 1904, name: 'Amara Okafor', email: 'amara@kinetic.example', phone: null, job_title: 'Head of Partnerships', company_name: 'Kinetic Works', company_domain: 'kinetic.example', linkedin_url: 'https://linkedin.com', country: 'CA', why_hot: 'Happy to take a look. Please send the one-pager and I will loop in our sales lead.', became_hot_at: '2026-09-09T12:41:00Z', status: 'New' },
  { person_id: 'demo-4', campaign_id: 1904, name: 'Leo Martins', email: 'leo@orbitpath.example', phone: '+351 21 555 0124', job_title: 'COO', company_name: 'OrbitPath', company_domain: 'orbitpath.example', linkedin_url: 'https://linkedin.com', country: 'PT', why_hot: 'Could be useful for us. Are you available Thursday afternoon?', became_hot_at: '2026-09-08T18:15:00Z', status: 'Sent' },
];
const defaultSettings: DraftSettings = {
  instruction: 'Be warm, direct, and useful. Keep replies under 100 words. End with one clear next step.',
  example: 'Hi {{first_name}},\n\nThanks for getting back to me — happy to share more. Would Tuesday at 10:00 or Wednesday at 14:00 work for a quick conversation?\n\nBest,\n{{signature}}',
  signature: 'Yash',
};
const apiBase = process.env.NEXT_PUBLIC_API_BASE_URL?.replace(/\/$/, '') || 'https://explee-signal-desk-api.yashjay2003.workers.dev';
const microsoftClientId = process.env.NEXT_PUBLIC_MICROSOFT_CLIENT_ID || 'f9e05546-b4fb-45a6-888c-bbb4593a03cf';
const microsoftTenantId = process.env.NEXT_PUBLIC_MICROSOFT_TENANT_ID || '50469384-abf0-4337-b9fe-6c48161ec38c';
let msal: PublicClientApplication | null = null;

function leadKey(lead: Lead) { return `${lead.campaign_id}:${lead.person_id}`; }
function firstName(name: string | null) { return name?.trim().split(/\s+/)[0] || 'there'; }
function makeDraft(lead: Lead, settings: DraftSettings) {
  const sample = settings.example.trim();
  if (sample) return sample.replaceAll('{{first_name}}', firstName(lead.name)).replaceAll('{{company_name}}', lead.company_name || 'your team').replaceAll('{{signature}}', settings.signature);
  return `Hi ${firstName(lead.name)},\n\nThanks for getting back to me — happy to share more. Would Tuesday or Wednesday work for a quick conversation?\n\nBest,\n${settings.signature}`;
}
function formatWhen(value: string | null) {
  if (!value) return 'Recently';
  const date = new Date(value), today = new Date();
  if (date.toDateString() === today.toDateString()) return date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  return date.toLocaleDateString([], { month: 'short', day: 'numeric' });
}
async function getMsal() {
  if (!microsoftClientId) throw new Error('Microsoft sign-in is not configured yet.');
  if (!msal) {
    msal = new PublicClientApplication({ auth: { clientId: microsoftClientId, authority: `https://login.microsoftonline.com/${microsoftTenantId}`, redirectUri: window.location.href.split('#')[0] }, cache: { cacheLocation: 'localStorage' } });
    await msal.initialize();
  }
  return msal;
}

export default function Home() {
  const [view, setView] = useState<'leads' | 'settings'>('leads');
  const [leads, setLeads] = useState<Lead[]>(demoLeads);
  const [selectedKey, setSelectedKey] = useState(leadKey(demoLeads[0]));
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(0);
  const [total, setTotal] = useState(demoLeads.length);
  const [settings, setSettings] = useState(defaultSettings);
  const [draft, setDraft] = useState(makeDraft(demoLeads[0], defaultSettings));
  const [subject, setSubject] = useState('Re: Next steps');
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [idToken, setIdToken] = useState('');
  const [graphToken, setGraphToken] = useState('');
  const [loading, setLoading] = useState(false);
  const [sending, setSending] = useState(false);
  const [notice, setNotice] = useState('');
  const [showAll, setShowAll] = useState(true);
  const selected = leads.find((lead) => leadKey(lead) === selectedKey) || leads[0];
  const filtered = useMemo(() => {
    const needle = query.toLowerCase().trim();
    return leads.filter((lead) => (showAll || lead.status !== 'Sent') && (!needle || [lead.name, lead.email, lead.company_name, lead.job_title].filter(Boolean).some((value) => value!.toLowerCase().includes(needle))));
  }, [leads, query, showAll]);
  const authedFetch = useCallback(async (path: string, init?: RequestInit) => {
    if (!apiBase || !idToken) throw new Error('Connect Microsoft 365 first.');
    const headers = new Headers(init?.headers);
    headers.set('Authorization', `Bearer ${idToken}`);
    headers.set('Content-Type', 'application/json');
    return fetch(`${apiBase}${path}`, { ...init, headers });
  }, [idToken]);
  const loadLeads = useCallback(async () => {
    if (!apiBase || !idToken) return;
    setLoading(true);
    try {
      const response = await authedFetch(`/api/leads?limit=100&offset=${page * 100}`);
      if (!response.ok) throw new Error('Could not load leads.');
      const payload = (await response.json()) as { leads: Lead[]; total: number };
      setLeads(payload.leads);
      setTotal(payload.total);
      if (payload.leads[0]) setSelectedKey((current) => payload.leads.some((lead) => leadKey(lead) === current) ? current : leadKey(payload.leads[0]));
      const settingsResponse = await authedFetch('/api/settings');
      if (settingsResponse.ok) setSettings(await settingsResponse.json());
    } catch (error) { setNotice(error instanceof Error ? error.message : 'Could not load leads.'); }
    finally { setLoading(false); }
  }, [authedFetch, idToken, page]);

  useEffect(() => { void (async () => {
    if (!microsoftClientId) return;
    const client = await getMsal(), active = client.getAllAccounts()[0];
    if (!active) return;
    try {
      const result = await client.acquireTokenSilent({ account: active, scopes: ['User.Read', 'Mail.Send'] });
      setAccount(active); setGraphToken(result.accessToken); setIdToken(result.idToken);
    } catch { /* The connect button resumes sign-in. */ }
  })(); }, []);
  useEffect(() => {
    if (!idToken) return;
    const pending = window.setTimeout(() => void loadLeads(), 0);
    return () => window.clearTimeout(pending);
  }, [idToken, loadLeads]);
  useEffect(() => {
    const context = (document as Document & { modelContext?: { registerTool?: (tool: unknown, options: unknown) => unknown } }).modelContext;
    if (!context?.registerTool) return;
    const lifecycle = new AbortController();
    void Promise.resolve(context.registerTool({ name: 'list_hot_leads', title: 'List hot leads', description: 'Read the hot leads currently shown in Signal Desk.', inputSchema: { type: 'object', properties: {}, additionalProperties: false }, annotations: { readOnlyHint: true, untrustedContentHint: true }, execute: () => ({ leads: filtered.map(({ name, email, phone, company_name, why_hot }) => ({ name, email, phone, company_name, why_hot })) }) }, { signal: lifecycle.signal })).catch(() => undefined);
    return () => lifecycle.abort();
  }, [filtered]);
  useEffect(() => {
    if (!idToken) return;
    const importBatch = async () => {
      const node = document.getElementById('signal-desk-import-data');
      if (!node?.textContent) return;
      try {
        const prospects = JSON.parse(node.textContent) as Lead[];
        const response = await authedFetch('/api/prospects/import', { method: 'POST', body: JSON.stringify({ prospects }) });
        if (!response.ok) throw new Error(`Import failed (${response.status}).`);
        node.dataset.status = 'complete';
      } catch (error) {
        node.dataset.status = error instanceof Error ? error.message : 'Import failed.';
      }
    };
    document.addEventListener('signal-desk-import', importBatch);
    return () => document.removeEventListener('signal-desk-import', importBatch);
  }, [authedFetch, idToken]);

  async function connectMicrosoft() {
    try {
      const client = await getMsal(), result = await client.loginPopup({ scopes: ['User.Read', 'Mail.Send'] });
      setAccount(result.account); setGraphToken(result.accessToken); setIdToken(result.idToken); setNotice('Microsoft 365 connected.');
    } catch (error) { setNotice(error instanceof Error ? error.message : 'Sign-in was not completed.'); }
  }
  function chooseLead(lead: Lead) {
    setSelectedKey(leadKey(lead));
    setDraft(makeDraft(lead, settings));
  }
  async function saveSettings() {
    if (!apiBase || !idToken) return setNotice('Settings saved for this preview.');
    const response = await authedFetch('/api/settings', { method: 'PUT', body: JSON.stringify(settings) });
    setNotice(response.ok ? 'Draft settings saved.' : 'Could not save settings.');
  }
  async function syncNow() {
    if (!apiBase || !idToken) return;
    setLoading(true);
    try {
      const response = await authedFetch('/api/sync', { method: 'POST' });
      if (!response.ok) throw new Error('Explee sync did not complete.');
      await loadLeads();
      setNotice('Enterprise HR leads are up to date.');
    } catch (error) { setNotice(error instanceof Error ? error.message : 'Could not sync leads.'); }
    finally { setLoading(false); }
  }
  async function sendMail() {
    if (!selected?.email) return setNotice('This lead does not have an email address.');
    if (!graphToken) return void connectMicrosoft();
    setSending(true);
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', { method: 'POST', headers: { Authorization: `Bearer ${graphToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ message: { subject, body: { contentType: 'Text', content: draft }, toRecipients: [{ emailAddress: { address: selected.email, name: selected.name || undefined } }] }, saveToSentItems: true }) });
      if (!response.ok) throw new Error('Outlook did not send the message. Please reconnect and try again.');
      setLeads((items) => items.map((lead) => leadKey(lead) === selectedKey ? { ...lead, status: 'Sent' } : lead));
      if (apiBase && idToken) void authedFetch(`/api/leads/${selected.campaign_id}/${encodeURIComponent(selected.person_id)}/status`, { method: 'PUT', body: JSON.stringify({ status: 'Sent' }) });
      setNotice(`Email sent to ${selected.email}.`);
    } catch (error) { setNotice(error instanceof Error ? error.message : 'Email could not be sent.'); }
    finally { setSending(false); }
  }

  return <main className="app-shell">
    <aside className="sidebar">
      <div className="brand-mark"><Flame size={19} strokeWidth={2.4} /></div>
      <nav aria-label="Main navigation"><button className={view === 'leads' ? 'nav-item active' : 'nav-item'} onClick={() => setView('leads')} aria-label="Enterprise HR leads"><UserRound size={19} /></button><button className={view === 'settings' ? 'nav-item active' : 'nav-item'} onClick={() => setView('settings')} aria-label="Settings"><Settings2 size={19} /></button></nav>
      <div className="sidebar-spacer" /><div className="avatar">{account?.name?.[0] || 'Y'}</div>
    </aside>
    <section className="workspace">
      <header className="topbar"><div><div className="eyebrow"><span className="live-dot" /> EXPLEE · ENTERPRISE HR</div><h1>{view === 'leads' ? 'Enterprise HR leads' : 'Draft settings'}</h1></div><div className="top-actions"><div className="sync-copy"><strong>{apiBase ? 'Hot-lead sync on' : 'Preview mode'}</strong><span>{apiBase ? 'Every 10 minutes' : 'Connect backend to go live'}</span></div><Button variant="outline" className="quiet-button" onClick={() => void syncNow()} disabled={loading || !idToken}><RefreshCw size={15} className={loading ? 'spin' : ''} /> Sync now</Button><Button className="connect-button" onClick={() => void connectMicrosoft()}><span className="microsoft-mark"><i /><i /><i /><i /></span>{account ? account.username : 'Connect Microsoft 365'}</Button></div></header>
      {notice && <output className="notice"><Check size={15} /> {notice}<button onClick={() => setNotice('')} aria-label="Dismiss"><X size={14} /></button></output>}
      {view === 'settings' ? <section className="settings-page">
        <div className="settings-intro"><span className="settings-icon"><Sparkles size={20} /></span><div><h2>Shape every first draft</h2><p>Signal Desk uses your example as a reusable reply template. You can always edit the result before sending.</p></div></div>
        <div className="settings-card"><div className="field-block"><Label htmlFor="instruction">Writing instruction</Label><p>Describe the voice and rules you want replies to follow.</p><Textarea id="instruction" value={settings.instruction} onChange={(event) => setSettings({ ...settings, instruction: event.target.value })} rows={4} /></div><div className="field-block"><Label htmlFor="example">Reply example</Label><p>Use <code>{'{{first_name}}'}</code>, <code>{'{{company_name}}'}</code>, and <code>{'{{signature}}'}</code> as placeholders.</p><Textarea id="example" value={settings.example} onChange={(event) => setSettings({ ...settings, example: event.target.value })} rows={8} /></div><div className="field-block narrow"><Label htmlFor="signature">Signature name</Label><Input id="signature" value={settings.signature} onChange={(event) => setSettings({ ...settings, signature: event.target.value })} /></div><div className="settings-footer"><Button className="primary-button" onClick={() => void saveSettings()}>Save settings</Button></div></div>
      </section> : <div className="lead-layout">
        <section className="lead-list"><div className="list-toolbar"><div className="search-wrap"><Search size={16} /><Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search this page…" aria-label="Search leads" /></div><button className="filter-button" onClick={() => setShowAll((value) => !value)}><SlidersHorizontal size={16} /> {showAll ? 'All leads' : 'Open only'} <ChevronDown size={14} /></button><button className="page-button" disabled={page === 0} onClick={() => setPage((value) => Math.max(0, value - 1))}>Previous</button><button className="page-button" disabled={(page + 1) * 100 >= total} onClick={() => setPage((value) => value + 1)}>Next</button><span className="result-count">{page * 100 + 1}–{Math.min((page + 1) * 100, total)} of {total}</span></div><div className="table-wrap"><Table><TableHeader><TableRow><TableHead className="name-col">LEAD</TableHead><TableHead>COMPANY</TableHead><TableHead>CONTACT</TableHead><TableHead>DETAILS</TableHead><TableHead>STATUS</TableHead><TableHead className="time-col">ADDED</TableHead></TableRow></TableHeader><TableBody>
          {filtered.map((lead) => <TableRow key={leadKey(lead)} data-active={leadKey(lead) === selectedKey} onClick={() => chooseLead(lead)} tabIndex={0} onKeyDown={(event) => { if (event.key === 'Enter') chooseLead(lead); }}><TableCell><div className="person-cell"><span className="person-avatar">{lead.name?.split(' ').map((part) => part[0]).join('').slice(0, 2) || '?'}</span><span><strong>{lead.name || 'Unknown lead'}</strong><small>{lead.job_title || '—'}</small></span></div></TableCell><TableCell><strong className="company-name">{lead.company_name || '—'}</strong><small className="cell-meta">{lead.country || ''}</small></TableCell><TableCell><span className="contact-line"><AtSign size={13} />{lead.email || 'Not enriched'}</span><span className="contact-line muted"><Phone size={13} />{lead.phone || 'Not enriched'}</span></TableCell><TableCell><span className="interest-preview">{lead.why_hot ? `“${lead.why_hot}”` : 'Enterprise HR campaign prospect'}</span></TableCell><TableCell><span className={`status status-${(lead.status || 'New').toLowerCase()}`}>{lead.status || 'New'}</span></TableCell><TableCell><span className="when"><Clock3 size={13} />{lead.became_hot_at ? formatWhen(lead.became_hot_at) : 'Sourced'}</span></TableCell></TableRow>)}
        </TableBody></Table></div></section>
        {selected && <aside className="composer" aria-label={`Lead details for ${selected.name}`}><div className="composer-head"><div><span className="hot-label"><Flame size={13} /> {selected.why_hot ? 'HOT LEAD' : 'PROSPECT'}</span><h2>{selected.name || 'Unknown lead'}</h2><p>{selected.job_title || 'Contact'} at {selected.company_name || 'Unknown company'}</p></div>{selected.linkedin_url && <a href={selected.linkedin_url} target="_blank" rel="noreferrer" aria-label="Open LinkedIn"><ArrowUpRight size={17} /></a>}</div><div className="verified-contact"><div><Mail size={16} /><span><small>EMAIL</small><strong>{selected.email || 'Not enriched by Explee'}</strong></span></div><div><Phone size={16} /><span><small>PHONE</small><strong>{selected.phone || 'Not enriched by Explee'}</strong></span></div></div><div className="reply-card"><div className="reply-label">{selected.why_hot ? <>THEIR REPLY <span>{formatWhen(selected.became_hot_at)}</span></> : 'SOURCE'}</div><blockquote>{selected.why_hot ? `“${selected.why_hot}”` : 'Enterprise HR teams campaign'}</blockquote></div><div className="draft-head"><div><FilePenLine size={16} /><strong>Your draft</strong><span>Editable</span></div><button onClick={() => setDraft(makeDraft(selected, settings))}><Sparkles size={14} /> Regenerate</button></div><div className="mail-field"><span>TO</span><strong>{selected.email || 'Email not available yet'}</strong></div><div className="mail-field"><span>SUBJECT</span><Input value={subject} onChange={(event) => setSubject(event.target.value)} aria-label="Email subject" /></div><Textarea className="draft-area" value={draft} onChange={(event) => setDraft(event.target.value)} aria-label="Email draft" /><div className="send-footer"><div>{account ? <><span className="connected-dot" />Sending from <strong>{account.username}</strong></> : <><CircleAlert size={14} /> Connect Outlook before sending</>}</div><Button className="send-button" onClick={() => void sendMail()} disabled={sending || !selected.email}><Send size={15} /> {sending ? 'Sending…' : 'Review & send'}</Button></div></aside>}
      </div>}
    </section>
  </main>;
}
