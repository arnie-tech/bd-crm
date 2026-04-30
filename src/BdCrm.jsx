import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const SB_URL = import.meta.env.VITE_SB_URL;
const SB_KEY = import.meta.env.VITE_SB_KEY;

const H = {
  apikey: SB_KEY,
  Authorization: `Bearer ${SB_KEY}`,
  "Content-Type": "application/json",
  Prefer: "return=representation",
};

const api = async (path, options = {}) => {
  const res = await fetch(`${SB_URL}/rest/v1/${path}`, { headers: H, ...options });
  if (!res.ok) throw new Error(await res.text());
  const text = await res.text();
  return text ? JSON.parse(text) : [];
};

const ACTIVITY_TYPES = ["Phone Call", "Email", "LinkedIn", "Meeting", "Seminar", "Direct Mail", "Referral Received"];
const CONTACT_TYPES = ["Lawyer", "Barrister"];
const FIRM_TYPES = ["Law Firm", "Barrister Chambers"];
const AU_STATES = ["NSW", "VIC", "QLD", "WA", "SA", "TAS", "ACT", "NT"];

const getMailingAddress = (c) => {
  const hasOverride = c.mailing_line1 || c.mailing_suburb || c.mailing_postcode;
  const src = hasOverride
    ? { line1: c.mailing_line1, line2: c.mailing_line2, suburb: c.mailing_suburb, state: c.mailing_state, postcode: c.mailing_postcode }
    : c.firm
    ? { line1: c.firm.address_line1, line2: c.firm.address_line2, suburb: c.firm.suburb, state: c.firm.state, postcode: c.firm.postcode }
    : null;
  if (!src || !(src.line1 || src.suburb || src.postcode)) return null;
  return { ...src, source: hasOverride ? "contact" : "firm" };
};

const formatAddressLines = (a) => {
  if (!a) return [];
  const lines = [];
  if (a.line1) lines.push(a.line1);
  if (a.line2) lines.push(a.line2);
  const cityLine = [a.suburb, a.state, a.postcode].filter(Boolean).join(" ");
  if (cityLine) lines.push(cityLine);
  return lines;
};

const todayStr = () => new Date().toISOString().split("T")[0];
const fmtDate = (d) =>
  d ? new Date(d + "T00:00:00").toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const isOverdue = (d) => d && d < todayStr();

const inputStyle = {
  width: "100%",
  padding: "7px 10px",
  border: "1px solid #d1d5db",
  borderRadius: 6,
  fontSize: 13,
  boxSizing: "border-box",
  outline: "none",
};

// ---- Shared UI ----

function Modal({ title, onClose, children, width = 560 }) {
  return (
    <div
      onClick={(e) => e.target === e.currentTarget && onClose()}
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}
    >
      <div style={{ background: "white", borderRadius: 10, padding: 24, width: "100%", maxWidth: width, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <div style={{ fontWeight: 700, fontSize: 16, color: "#1e3a5f" }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#94a3b8", lineHeight: 1 }}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#374151", marginBottom: 4 }}>{label}</label>
      {children}
    </div>
  );
}

function TagInput({ value = [], onChange, allTags, onCreateTag }) {
  const [input, setInput] = useState("");
  const [open, setOpen] = useState(false);
  const filtered = allTags.filter(
    (t) => t.name.toLowerCase().includes(input.toLowerCase()) && !value.find((v) => v.id === t.id)
  );
  const addTag = (tag) => { onChange([...value, tag]); setInput(""); setOpen(false); };
  const handleKey = async (e) => {
    if (e.key === "Enter" && input.trim()) {
      e.preventDefault();
      const ex = allTags.find((t) => t.name.toLowerCase() === input.trim().toLowerCase());
      if (ex) addTag(ex);
      else { const nt = await onCreateTag(input.trim()); if (nt) addTag(nt); }
    }
  };
  return (
    <div style={{ border: "1px solid #d1d5db", borderRadius: 6, padding: "4px 6px", minHeight: 36, display: "flex", flexWrap: "wrap", gap: 4, position: "relative" }}>
      {value.map((t) => (
        <span key={t.id} style={{ background: "#dbeafe", color: "#1d4ed8", borderRadius: 99, padding: "2px 8px", fontSize: 11, display: "flex", alignItems: "center", gap: 4 }}>
          {t.name}
          <button onClick={() => onChange(value.filter((v) => v.id !== t.id))} style={{ background: "none", border: "none", cursor: "pointer", color: "#1d4ed8", fontSize: 14, lineHeight: 1, padding: 0 }}>×</button>
        </span>
      ))}
      <div style={{ position: "relative" }}>
        <input
          value={input}
          onChange={(e) => { setInput(e.target.value); setOpen(true); }}
          onKeyDown={handleKey}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          placeholder="Add tag..."
          style={{ border: "none", outline: "none", fontSize: 13, minWidth: 80, padding: "2px 4px" }}
        />
        {open && (filtered.length > 0 || (input && !allTags.find((t) => t.name.toLowerCase() === input.toLowerCase()))) && (
          <div style={{ position: "absolute", top: "100%", left: 0, background: "white", border: "1px solid #e2e8f0", borderRadius: 6, boxShadow: "0 4px 12px rgba(0,0,0,0.1)", zIndex: 50, minWidth: 160, marginTop: 2 }}>
            {filtered.map((t) => (
              <div key={t.id} onMouseDown={() => addTag(t)} style={{ padding: "7px 12px", cursor: "pointer", fontSize: 13, color: "#374151" }}>{t.name}</div>
            ))}
            {input && !allTags.find((t) => t.name.toLowerCase() === input.toLowerCase()) && (
              <div
                onMouseDown={async () => { const nt = await onCreateTag(input.trim()); if (nt) addTag(nt); }}
                style={{ padding: "7px 12px", cursor: "pointer", fontSize: 13, color: "#2d7dd2", borderTop: "1px solid #f1f5f9" }}
              >
                + Create "{input}"
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ---- Pages ----

function Dashboard({ contacts, activities, firms, onNav, onLogActivity }) {
  const td = todayStr();
  const overdueItems = activities.filter((a) => a.follow_up_date && a.follow_up_date < td);
  const upcoming = [...activities]
    .filter((a) => a.follow_up_date && a.follow_up_date >= td)
    .sort((a, b) => a.follow_up_date.localeCompare(b.follow_up_date))
    .slice(0, 10);
  const recent = activities.slice(0, 15);
  const refCounts = {};
  activities.filter((a) => a.type === "Referral Received").forEach((a) => {
    refCounts[a.contact_id] = (refCounts[a.contact_id] || 0) + 1;
  });
  const topReferrers = Object.entries(refCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([id, count]) => ({ c: contacts.find((x) => x.id === id), count }))
    .filter((x) => x.c);

  const Card = ({ title, children }) => (
    <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16, marginBottom: 16 }}>
      <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 12 }}>{title}</div>
      {children}
    </div>
  );
  const Row = ({ children }) => (
    <div style={{ padding: "7px 0", borderBottom: "1px solid #f1f5f9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>{children}</div>
  );
  const ContactLink = ({ c }) => (
    <button
      onClick={() => onNav("contact-detail", { contactId: c.id })}
      style={{ background: "none", border: "none", cursor: "pointer", color: "#2d7dd2", fontWeight: 600, fontSize: 13, padding: 0, textAlign: "left" }}
    >
      {c.first_name} {c.last_name}
    </button>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Dashboard</h1>
        <button onClick={onLogActivity} style={{ padding: "8px 18px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>
          + Log Activity
        </button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div>
          {overdueItems.length > 0 && (
            <Card title={`⚠️ Overdue Follow-ups (${overdueItems.length})`}>
              {overdueItems.slice(0, 6).map((a) => {
                const c = contacts.find((x) => x.id === a.contact_id);
                return c ? (
                  <Row key={a.id}>
                    <ContactLink c={c} />
                    <span style={{ fontSize: 12, color: "#ef4444" }}>{fmtDate(a.follow_up_date)}</span>
                  </Row>
                ) : null;
              })}
            </Card>
          )}
          <Card title="📅 Upcoming Follow-ups">
            {upcoming.length === 0 ? (
              <p style={{ color: "#94a3b8", fontSize: 13, margin: 0 }}>No upcoming follow-ups.</p>
            ) : (
              upcoming.map((a) => {
                const c = contacts.find((x) => x.id === a.contact_id);
                return c ? (
                  <Row key={a.id}>
                    <div>
                      <ContactLink c={c} />
                      <div style={{ fontSize: 11, color: "#64748b" }}>{a.type}{a.notes ? ` · ${a.notes.slice(0, 35)}` : ""}</div>
                    </div>
                    <span style={{ fontSize: 12, color: "#64748b", whiteSpace: "nowrap", marginLeft: 8 }}>{fmtDate(a.follow_up_date)}</span>
                  </Row>
                ) : null;
              })
            )}
          </Card>
        </div>
        <div>
          <Card title="🏆 Top Referrers">
            {topReferrers.length === 0 ? (
              <p style={{ color: "#94a3b8", fontSize: 13, margin: 0 }}>No referrals logged yet.</p>
            ) : (
              topReferrers.map(({ c, count }) => (
                <Row key={c.id}>
                  <div>
                    <ContactLink c={c} />
                    <div style={{ fontSize: 11, color: "#94a3b8" }}>{c.firm?.name || ""}</div>
                  </div>
                  <span style={{ background: "#dcfce7", color: "#16a34a", borderRadius: 99, padding: "2px 10px", fontSize: 12, fontWeight: 700 }}>{count}</span>
                </Row>
              ))
            )}
          </Card>
          <Card title="📋 Recent Activity">
            {recent.length === 0 ? (
              <p style={{ color: "#94a3b8", fontSize: 13, margin: 0 }}>No activities yet.</p>
            ) : (
              recent.map((a) => {
                const c = contacts.find((x) => x.id === a.contact_id);
                const colors = { "Referral Received": "#16a34a", Meeting: "#7c3aed", "Phone Call": "#d97706" };
                return (
                  <div key={a.id} style={{ padding: "7px 0", borderBottom: "1px solid #f1f5f9" }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, fontWeight: 600, color: colors[a.type] || "#2d7dd2" }}>{a.type}</span>
                      <span style={{ fontSize: 11, color: "#94a3b8" }}>{fmtDate(a.date)}</span>
                    </div>
                    {c && (
                      <button onClick={() => onNav("contact-detail", { contactId: c.id })} style={{ background: "none", border: "none", cursor: "pointer", color: "#374151", fontSize: 13, padding: 0, fontWeight: 500 }}>
                        {c.first_name} {c.last_name}
                      </button>
                    )}
                    {a.notes && <div style={{ fontSize: 11, color: "#94a3b8" }}>{a.notes.slice(0, 55)}{a.notes.length > 55 ? "..." : ""}</div>}
                  </div>
                );
              })
            )}
          </Card>
        </div>
      </div>
    </div>
  );
}

function ContactsList({ contacts, firms, tags, onSelect, onAdd, onEdit }) {
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("All");
  const [tagFilter, setTagFilter] = useState("");
  const [firmFilter, setFirmFilter] = useState("");

  const filtered = contacts.filter((c) => {
    const s = search.toLowerCase();
    return (
      (!s || `${c.first_name} ${c.last_name} ${c.email || ""} ${c.firm?.name || ""}`.toLowerCase().includes(s)) &&
      (typeFilter === "All" || c.type === typeFilter) &&
      (!tagFilter || c.tags?.some((t) => t.id === tagFilter)) &&
      (!firmFilter || c.firm_id === firmFilter)
    );
  });

  const Th = ({ ch }) => (
    <th style={{ padding: "8px 12px", textAlign: "left", fontSize: 12, fontWeight: 600, color: "#64748b", borderBottom: "1px solid #e2e8f0", background: "#f8fafc" }}>{ch}</th>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Contacts ({filtered.length})</h1>
        <button onClick={onAdd} style={{ padding: "8px 18px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>
          + Add Contact
        </button>
      </div>
      <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 10, marginBottom: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
        <input
          placeholder="Search name, email, firm..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{ flex: 1, minWidth: 180, padding: "6px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13, outline: "none" }}
        />
        <select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} style={{ padding: "6px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13 }}>
          {["All", ...CONTACT_TYPES].map((o) => <option key={o}>{o}</option>)}
        </select>
        <select value={firmFilter} onChange={(e) => setFirmFilter(e.target.value)} style={{ padding: "6px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13 }}>
          <option value="">All Firms</option>
          {firms.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select>
        <select value={tagFilter} onChange={(e) => setTagFilter(e.target.value)} style={{ padding: "6px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13 }}>
          <option value="">All Tags</option>
          {tags.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
        </select>
      </div>
      <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>{["Name", "Type", "Firm", "Email", "Tags", ""].map((h) => <Th key={h} ch={h} />)}</tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={6} style={{ padding: 32, textAlign: "center", color: "#94a3b8", fontSize: 14 }}>No contacts found</td></tr>
            ) : (
              filtered.map((c) => (
                <tr key={c.id} style={{ borderTop: "1px solid #f1f5f9" }}>
                  <td style={{ padding: "10px 12px" }}>
                    <button onClick={() => onSelect(c)} style={{ background: "none", border: "none", cursor: "pointer", color: "#2d7dd2", fontWeight: 600, fontSize: 14, padding: 0 }}>
                      {c.first_name} {c.last_name}
                    </button>
                  </td>
                  <td style={{ padding: "10px 12px", fontSize: 13, color: "#64748b" }}>{c.type}</td>
                  <td style={{ padding: "10px 12px", fontSize: 13, color: "#64748b" }}>{c.firm?.name || "—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 13, color: "#64748b" }}>{c.email || "—"}</td>
                  <td style={{ padding: "10px 12px" }}>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
                      {(c.tags || []).map((t) => (
                        <span key={t.id} style={{ background: "#dbeafe", color: "#1d4ed8", borderRadius: 99, padding: "1px 8px", fontSize: 11 }}>{t.name}</span>
                      ))}
                    </div>
                  </td>
                  <td style={{ padding: "10px 12px" }}>
                    <button onClick={() => onEdit(c)} style={{ background: "none", border: "1px solid #d1d5db", borderRadius: 4, padding: "3px 10px", cursor: "pointer", fontSize: 12, color: "#374151" }}>Edit</button>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ContactDetail({ contact: c, activities, onEdit, onDelete, onLog, onBack }) {
  const referralCount = activities.filter((a) => a.type === "Referral Received").length;
  const typeColors = {
    "Referral Received": "#16a34a", Meeting: "#7c3aed", "Phone Call": "#d97706",
    Email: "#0891b2", LinkedIn: "#0369a1", Seminar: "#9333ea", "Direct Mail": "#be185d",
  };

  return (
    <div>
      <button onClick={onBack} style={{ background: "none", border: "none", cursor: "pointer", color: "#64748b", fontSize: 13, marginBottom: 12, padding: 0 }}>
        ← Back to Contacts
      </button>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
        <div>
          <h1 style={{ fontSize: 24, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>{c.first_name} {c.last_name}</h1>
          <div style={{ color: "#64748b", fontSize: 14, marginTop: 4 }}>{c.type}{c.firm ? ` · ${c.firm.name}` : ""}</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onLog} style={{ padding: "8px 16px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>+ Log Activity</button>
          <button onClick={onEdit} style={{ padding: "8px 14px", background: "white", border: "1px solid #d1d5db", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Edit</button>
          <button onClick={onDelete} style={{ padding: "8px 14px", background: "white", border: "1px solid #fca5a5", borderRadius: 6, cursor: "pointer", color: "#ef4444", fontSize: 13 }}>Delete</button>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "300px 1fr", gap: 16 }}>
        <div>
          <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16, marginBottom: 12 }}>
            <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 10 }}>Contact Info</div>
            {[
              ["Email", c.email ? <a href={`mailto:${c.email}`} style={{ color: "#2d7dd2", textDecoration: "none" }}>{c.email}</a> : "—"],
              ["Phone", c.phone || "—"],
              ["LinkedIn", c.linkedin_url ? <a href={c.linkedin_url} target="_blank" rel="noreferrer" style={{ color: "#2d7dd2", textDecoration: "none" }}>View Profile</a> : "—"],
              ["Firm", c.firm?.name || "—"],
              ["Referrals", <span style={{ background: "#dcfce7", color: "#16a34a", borderRadius: 99, padding: "1px 10px", fontWeight: 700, fontSize: 13 }}>{referralCount}</span>],
            ].map(([label, val]) => (
              <div key={label} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: "1px solid #f8fafc", fontSize: 13 }}>
                <span style={{ color: "#64748b" }}>{label}</span>
                <span>{val}</span>
              </div>
            ))}
          </div>
          {(() => {
            const addr = getMailingAddress(c);
            const lines = formatAddressLines(addr);
            return (
              <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16, marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ fontWeight: 600, fontSize: 14, color: "#374151" }}>Mailing Address</div>
                  {addr && (
                    <span style={{ background: addr.source === "contact" ? "#fef3c7" : "#e0e7ff", color: addr.source === "contact" ? "#92400e" : "#3730a3", borderRadius: 99, padding: "2px 8px", fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4 }}>
                      {addr.source === "contact" ? "Contact override" : "From firm"}
                    </span>
                  )}
                </div>
                {lines.length === 0 ? (
                  <div style={{ fontSize: 13, color: "#94a3b8" }}>No address on file.</div>
                ) : (
                  <div style={{ fontSize: 13, color: "#374151", lineHeight: 1.5 }}>
                    {lines.map((l, i) => <div key={i}>{l}</div>)}
                  </div>
                )}
              </div>
            );
          })()}
          {c.tags?.length > 0 && (
            <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16, marginBottom: 12 }}>
              <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 8 }}>Tags</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {c.tags.map((t) => <span key={t.id} style={{ background: "#dbeafe", color: "#1d4ed8", borderRadius: 99, padding: "3px 10px", fontSize: 12 }}>{t.name}</span>)}
              </div>
            </div>
          )}
          {c.notes && (
            <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16 }}>
              <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 6 }}>Notes</div>
              <div style={{ fontSize: 13, color: "#374151", whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{c.notes}</div>
            </div>
          )}
        </div>
        <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 16 }}>
          <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 12 }}>Activity History ({activities.length})</div>
          {activities.length === 0 ? (
            <p style={{ color: "#94a3b8", fontSize: 13 }}>No activities logged yet.</p>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {activities.map((a) => (
                <div key={a.id} style={{ padding: 12, background: "#f8fafc", borderRadius: 6, borderLeft: `3px solid ${typeColors[a.type] || "#2d7dd2"}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <span style={{ fontWeight: 600, fontSize: 13, color: "#374151" }}>{a.type}</span>
                      {a.notes && <div style={{ fontSize: 12, color: "#64748b", marginTop: 2, lineHeight: 1.4 }}>{a.notes}</div>}
                    </div>
                    <div style={{ textAlign: "right", minWidth: 90 }}>
                      <div style={{ fontSize: 12, color: "#64748b" }}>{fmtDate(a.date)}</div>
                      {a.follow_up_date && (
                        <div style={{ fontSize: 11, color: isOverdue(a.follow_up_date) ? "#ef4444" : "#16a34a", marginTop: 2 }}>↻ {fmtDate(a.follow_up_date)}</div>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function FirmsList({ firms, contacts, onAdd, onEdit, onSelectContact }) {
  const [search, setSearch] = useState("");
  const [expanded, setExpanded] = useState(null);
  const filtered = firms.filter((f) => f.name.toLowerCase().includes(search.toLowerCase()));

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e3a5f", margin: 0 }}>Firms ({filtered.length})</h1>
        <button onClick={onAdd} style={{ padding: "8px 18px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 13 }}>
          + Add Firm
        </button>
      </div>
      <input placeholder="Search firms..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ ...inputStyle, marginBottom: 10 }} />
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {filtered.map((f) => {
          const fc = contacts.filter((c) => c.firm_id === f.id);
          const open = expanded === f.id;
          return (
            <div key={f.id} style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8 }}>
              <div onClick={() => setExpanded(open ? null : f.id)} style={{ padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                <div>
                  <div style={{ fontWeight: 600, color: "#1e3a5f", fontSize: 14 }}>{f.name}</div>
                  <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>{f.type} · {fc.length} contact{fc.length !== 1 ? "s" : ""}</div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <button onClick={(e) => { e.stopPropagation(); onEdit(f); }} style={{ background: "none", border: "1px solid #d1d5db", borderRadius: 4, padding: "3px 10px", cursor: "pointer", fontSize: 12, color: "#374151" }}>Edit</button>
                  <span style={{ color: "#94a3b8", fontSize: 12 }}>{open ? "▲" : "▼"}</span>
                </div>
              </div>
              {open && fc.length > 0 && (
                <div style={{ borderTop: "1px solid #f1f5f9", padding: "8px 16px 12px" }}>
                  {fc.map((c) => (
                    <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "3px 0" }}>
                      <button onClick={() => onSelectContact(c)} style={{ background: "none", border: "none", cursor: "pointer", color: "#2d7dd2", fontWeight: 500, fontSize: 13, padding: 0 }}>
                        {c.first_name} {c.last_name}
                      </button>
                      <span style={{ fontSize: 11, color: "#94a3b8" }}>{c.type}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ImportPage({ firms, tags, onCreateTag, onImportComplete }) {
  const [step, setStep] = useState(1);
  const [rows, setRows] = useState([]);
  const [hdrs, setHdrs] = useState([]);
  const [mapping, setMapping] = useState({});
  const [batchTags, setBatchTags] = useState([]);
  const [batchType, setBatchType] = useState("Lawyer");
  const [result, setResult] = useState(null);
  const [importing, setImporting] = useState(false);

  const CRM_FIELDS = [
    { key: "first_name", label: "First Name *" },
    { key: "last_name", label: "Last Name *" },
    { key: "email", label: "Email" },
    { key: "phone", label: "Phone" },
    { key: "linkedin_url", label: "LinkedIn URL" },
    { key: "firm_name", label: "Firm Name" },
    { key: "address_line1", label: "Firm Address Line 1" },
    { key: "address_line2", label: "Firm Address Line 2" },
    { key: "suburb", label: "Firm Suburb" },
    { key: "state", label: "Firm State" },
    { key: "postcode", label: "Firm Postcode" },
    { key: "notes", label: "Notes" },
    { key: "_skip", label: "(Skip)" },
  ];
  const FIRM_ADDR_FIELDS = ["address_line1", "address_line2", "suburb", "state", "postcode"];

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      if (data.length < 2) return;
      const h = data[0].map(String);
      setHdrs(h);
      setRows(data.slice(1).filter((r) => r.some((x) => x)));
      const auto = {};
      h.forEach((hh, i) => {
        const l = hh.toLowerCase().replace(/[\s_-]/g, "");
        if (l.includes("first") || l === "firstname") auto[i] = "first_name";
        else if (l.includes("last") || l === "lastname") auto[i] = "last_name";
        else if (l.includes("email") || l === "emailaddress") auto[i] = "email";
        else if (l.includes("phone") || l.includes("mobile")) auto[i] = "phone";
        else if (l.includes("linkedin")) auto[i] = "linkedin_url";
        else if (l.includes("firm") || l.includes("company") || l.includes("org")) auto[i] = "firm_name";
        else if (l.includes("postcode") || l.includes("postalcode") || l === "zip") auto[i] = "postcode";
        else if (l === "state" || l.includes("state")) auto[i] = "state";
        else if (l.includes("suburb") || l === "city" || l === "town") auto[i] = "suburb";
        else if (l.includes("address2") || l.includes("addressline2") || l === "line2") auto[i] = "address_line2";
        else if (l.includes("address") || l === "street" || l === "line1") auto[i] = "address_line1";
        else if (l.includes("note")) auto[i] = "notes";
        else auto[i] = "_skip";
      });
      setMapping(auto);
      setStep(2);
    };
    reader.readAsBinaryString(file);
  };

  const preview = rows.slice(0, 5).map((row) => {
    const obj = {};
    hdrs.forEach((h, i) => { const f = mapping[i]; if (f && f !== "_skip") obj[f] = row[i] || ""; });
    return obj;
  });

  const doImport = async () => {
    setImporting(true);
    let imported = 0, dupes = 0, errors = 0;
    let firstError = null;
    const existing = await api("crm_contacts?select=email");
    const existingEmails = new Set(existing.map((e) => e.email?.toLowerCase()).filter(Boolean));
    const firmCache = {};
    firms.forEach((f) => {
      const hasAddr = !!(f.address_line1 || f.suburb || f.postcode);
      firmCache[f.name.toLowerCase()] = { id: f.id, hasAddr };
    });

    for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
      const row = rows[rowIdx];
      const obj = {};
      hdrs.forEach((h, i) => { const f = mapping[i]; if (f && f !== "_skip") obj[f] = row[i] ? String(row[i]).trim() : ""; });
      if (!obj.first_name || !obj.last_name) {
        errors++;
        if (!firstError) firstError = `Row ${rowIdx + 2}: missing first or last name`;
        continue;
      }
      if (obj.email && existingEmails.has(obj.email.toLowerCase())) { dupes++; continue; }
      try {
        const rowAddr = {};
        FIRM_ADDR_FIELDS.forEach((k) => { if (obj[k]) rowAddr[k] = obj[k]; });
        const hasRowAddr = Object.keys(rowAddr).length > 0;
        let firm_id = null;
        if (obj.firm_name) {
          const fk = obj.firm_name.toLowerCase();
          if (firmCache[fk]) {
            firm_id = firmCache[fk].id;
            if (hasRowAddr && !firmCache[fk].hasAddr) {
              await api(`crm_firms?id=eq.${firm_id}`, { method: "PATCH", body: JSON.stringify(rowAddr) });
              firmCache[fk].hasAddr = true;
            }
          } else {
            const [nf] = await api("crm_firms", { method: "POST", body: JSON.stringify({ name: obj.firm_name, type: "Law Firm", ...rowAddr }) });
            firmCache[fk] = { id: nf.id, hasAddr: hasRowAddr };
            firm_id = nf.id;
          }
        }
        const { firm_name, address_line1, address_line2, suburb, state, postcode, ...cd } = obj;
        const [nc] = await api("crm_contacts", { method: "POST", body: JSON.stringify({ ...cd, type: batchType, firm_id }) });
        if (batchTags.length > 0) {
          await api("crm_contact_tags", { method: "POST", body: JSON.stringify(batchTags.map((t) => ({ contact_id: nc.id, tag_id: t.id }))) });
        }
        if (obj.email) existingEmails.add(obj.email.toLowerCase());
        imported++;
      } catch (e) {
        errors++;
        const msg = (e && e.message) ? e.message : String(e);
        console.error(`Import row ${rowIdx + 2} failed:`, msg, obj);
        if (!firstError) firstError = `Row ${rowIdx + 2}: ${msg}`;
      }
    }
    setResult({ imported, dupes, errors, firstError });
    setStep(4);
    setImporting(false);
    onImportComplete();
  };

  const stepLabels = ["Upload", "Map Columns", "Preview", "Done"];

  return (
    <div>
      <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e3a5f", marginBottom: 20 }}>Import Contacts</h1>
      <div style={{ display: "flex", marginBottom: 24 }}>
        {stepLabels.map((s, i) => (
          <div key={s} style={{ display: "flex", alignItems: "center" }}>
            <div style={{ padding: "6px 18px", background: step > i + 1 ? "#16a34a" : step === i + 1 ? "#2d7dd2" : "#e2e8f0", color: step >= i + 1 ? "white" : "#94a3b8", fontSize: 13, fontWeight: 500, borderRadius: i === 0 ? "6px 0 0 6px" : i === 3 ? "0 6px 6px 0" : 0 }}>
              {s}
            </div>
            {i < 3 && <div style={{ width: 1, background: "#e2e8f0" }} />}
          </div>
        ))}
      </div>

      {step === 1 && (
        <div style={{ background: "white", border: "2px dashed #d1d5db", borderRadius: 8, padding: 48, textAlign: "center" }}>
          <div style={{ fontSize: 40, marginBottom: 10 }}>📂</div>
          <div style={{ fontSize: 16, fontWeight: 600, color: "#374151", marginBottom: 6 }}>Upload Excel File</div>
          <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 20 }}>Supports .xlsx and .xls files. First row should be column headers.</div>
          <input type="file" accept=".xlsx,.xls" onChange={handleFile} id="xlFile" style={{ display: "none" }} />
          <label htmlFor="xlFile" style={{ padding: "10px 28px", background: "#2d7dd2", color: "white", borderRadius: 6, cursor: "pointer", fontWeight: 600, fontSize: 14 }}>Choose File</label>
        </div>
      )}

      {step === 2 && (
        <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 20 }}>
          <div style={{ fontWeight: 600, color: "#374151", marginBottom: 4 }}>Map Columns — {rows.length} rows detected</div>
          <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16 }}>Auto-mapped where possible. Adjust any that need changing.</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 20 }}>
            {hdrs.map((h, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 13, color: "#374151", minWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 500 }}>{h}</span>
                <span style={{ color: "#94a3b8", fontSize: 12 }}>→</span>
                <select value={mapping[i] || "_skip"} onChange={(e) => setMapping((m) => ({ ...m, [i]: e.target.value }))} style={{ flex: 1, padding: "5px 8px", border: "1px solid #d1d5db", borderRadius: 4, fontSize: 12 }}>
                  {CRM_FIELDS.map((f) => <option key={f.key} value={f.key}>{f.label}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{ borderTop: "1px solid #f1f5f9", paddingTop: 16, marginBottom: 16 }}>
            <div style={{ fontWeight: 600, fontSize: 14, color: "#374151", marginBottom: 10 }}>Batch Settings</div>
            <div style={{ display: "flex", gap: 16, alignItems: "flex-start" }}>
              <div>
                <label style={{ fontSize: 12, color: "#64748b", display: "block", marginBottom: 4 }}>Contact Type for all rows</label>
                <select value={batchType} onChange={(e) => setBatchType(e.target.value)} style={{ padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13 }}>
                  {CONTACT_TYPES.map((t) => <option key={t}>{t}</option>)}
                </select>
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ fontSize: 12, color: "#64748b", display: "block", marginBottom: 4 }}>Apply Tags to All Imported Contacts</label>
                <TagInput value={batchTags} onChange={setBatchTags} allTags={tags} onCreateTag={onCreateTag} />
              </div>
            </div>
          </div>
          <button onClick={() => setStep(3)} style={{ padding: "8px 20px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>Preview →</button>
        </div>
      )}

      {step === 3 && (
        <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 20 }}>
          <div style={{ fontWeight: 600, color: "#374151", marginBottom: 4 }}>Preview (first 5 of {rows.length} rows)</div>
          <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14 }}>Confirm the mapping looks correct before importing.</div>
          <div style={{ overflowX: "auto", marginBottom: 20 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#f8fafc" }}>
                  {Object.keys(preview[0] || {}).map((k) => (
                    <th key={k} style={{ padding: "6px 10px", textAlign: "left", fontWeight: 600, color: "#64748b", borderBottom: "1px solid #e2e8f0" }}>{k}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((row, i) => (
                  <tr key={i} style={{ borderTop: "1px solid #f1f5f9" }}>
                    {Object.values(row).map((v, j) => <td key={j} style={{ padding: "6px 10px", color: "#374151" }}>{v || "—"}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setStep(2)} style={{ padding: "8px 16px", background: "white", border: "1px solid #d1d5db", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>← Back</button>
            <button onClick={doImport} disabled={importing} style={{ padding: "8px 24px", background: importing ? "#94a3b8" : "#16a34a", color: "white", border: "none", borderRadius: 6, cursor: importing ? "not-allowed" : "pointer", fontWeight: 600, fontSize: 13 }}>
              {importing ? "Importing…" : `✓ Import ${rows.length} Contacts`}
            </button>
          </div>
        </div>
      )}

      {step === 4 && result && (
        <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 40, textAlign: "center" }}>
          <div style={{ fontSize: 48, marginBottom: 12 }}>{result.imported > 0 ? "✅" : "⚠️"}</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#374151", marginBottom: 20 }}>Import Complete</div>
          <div style={{ display: "flex", gap: 16, justifyContent: "center", marginBottom: 20 }}>
            {[["Imported", result.imported, "#16a34a", "#dcfce7"], ["Duplicates Skipped", result.dupes, "#d97706", "#fef9c3"], ["Errors", result.errors, "#dc2626", "#fee2e2"]].map(([l, v, c, bg]) => (
              <div key={l} style={{ background: bg, borderRadius: 8, padding: "14px 24px", minWidth: 110 }}>
                <div style={{ fontSize: 28, fontWeight: 700, color: c }}>{v}</div>
                <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>{l}</div>
              </div>
            ))}
          </div>
          {result.firstError && (
            <div style={{ background: "#fee2e2", border: "1px solid #fca5a5", borderRadius: 8, padding: "10px 14px", marginBottom: 24, color: "#991b1b", fontSize: 12, textAlign: "left", whiteSpace: "pre-wrap", maxWidth: 720, marginLeft: "auto", marginRight: "auto" }}>
              <strong>First error:</strong> {result.firstError}
              <div style={{ color: "#7f1d1d", marginTop: 6, fontSize: 11 }}>Open the browser console (F12) for full per-row details.</div>
            </div>
          )}
          <button
            onClick={() => { setStep(1); setRows([]); setHdrs([]); setMapping({}); setBatchTags([]); setResult(null); }}
            style={{ padding: "9px 24px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}
          >
            Import Another File
          </button>
        </div>
      )}
    </div>
  );
}

function SettingsPage({ tags, onRefresh }) {
  const [newTag, setNewTag] = useState("");
  const addTag = async () => {
    if (!newTag.trim()) return;
    try { await api("crm_tags", { method: "POST", body: JSON.stringify({ name: newTag.trim() }) }); setNewTag(""); onRefresh(); } catch (e) {}
  };
  const deleteTag = async (id) => {
    if (!confirm("Delete this tag? It will be removed from all contacts.")) return;
    await api(`crm_tags?id=eq.${id}`, { method: "DELETE" }); onRefresh();
  };
  return (
    <div>
      <h1 style={{ fontSize: 22, fontWeight: 700, color: "#1e3a5f", marginBottom: 20 }}>Settings</h1>
      <div style={{ background: "white", border: "1px solid #e2e8f0", borderRadius: 8, padding: 20, maxWidth: 440 }}>
        <div style={{ fontWeight: 600, color: "#374151", marginBottom: 12 }}>Manage Tags</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <input
            value={newTag} onChange={(e) => setNewTag(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addTag()}
            placeholder="New tag name..." style={{ flex: 1, padding: "7px 10px", border: "1px solid #d1d5db", borderRadius: 6, fontSize: 13, outline: "none" }}
          />
          <button onClick={addTag} style={{ padding: "7px 18px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>Add</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          {tags.length === 0 ? (
            <p style={{ color: "#94a3b8", fontSize: 13, margin: 0 }}>No tags yet.</p>
          ) : (
            tags.map((t) => (
              <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 12px", background: "#f8fafc", borderRadius: 6 }}>
                <span style={{ fontSize: 13, color: "#374151" }}>{t.name}</span>
                <button onClick={() => deleteTag(t.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ef4444", fontSize: 12 }}>Delete</button>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
}

// ---- Modals ----

function ContactModal({ contact, firms, tags, onCreateTag, onSave, onClose }) {
  const [form, setForm] = useState(contact || { type: "Lawyer", first_name: "", last_name: "", email: "", phone: "", linkedin_url: "", firm_id: "", notes: "", mailing_line1: "", mailing_line2: "", mailing_suburb: "", mailing_state: "", mailing_postcode: "" });
  const [selTags, setSelTags] = useState(contact?.tags || []);
  const [saving, setSaving] = useState(false);
  const [showMailing, setShowMailing] = useState(!!(contact?.mailing_line1 || contact?.mailing_suburb || contact?.mailing_postcode));
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const handleSave = async () => {
    if (!form.first_name || !form.last_name) return alert("First and last name are required.");
    setSaving(true); await onSave({ ...form, tags: selTags }); setSaving(false);
  };
  const clearMailing = () => {
    setForm((f) => ({ ...f, mailing_line1: "", mailing_line2: "", mailing_suburb: "", mailing_state: "", mailing_postcode: "" }));
    setShowMailing(false);
  };
  return (
    <Modal title={contact ? "Edit Contact" : "Add Contact"} onClose={onClose}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="First Name *"><input value={form.first_name} onChange={(e) => set("first_name", e.target.value)} style={inputStyle} /></Field>
        <Field label="Last Name *"><input value={form.last_name} onChange={(e) => set("last_name", e.target.value)} style={inputStyle} /></Field>
        <Field label="Type">
          <select value={form.type} onChange={(e) => set("type", e.target.value)} style={inputStyle}>
            {CONTACT_TYPES.map((t) => <option key={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Firm">
          <select value={form.firm_id || ""} onChange={(e) => set("firm_id", e.target.value || null)} style={inputStyle}>
            <option value="">— None —</option>
            {firms.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
          </select>
        </Field>
        <Field label="Email"><input type="email" value={form.email || ""} onChange={(e) => set("email", e.target.value)} style={inputStyle} /></Field>
        <Field label="Phone"><input value={form.phone || ""} onChange={(e) => set("phone", e.target.value)} style={inputStyle} /></Field>
      </div>
      <Field label="LinkedIn URL"><input value={form.linkedin_url || ""} onChange={(e) => set("linkedin_url", e.target.value)} style={inputStyle} /></Field>
      {!showMailing ? (
        <button type="button" onClick={() => setShowMailing(true)} style={{ background: "none", border: "1px dashed #cbd5e1", borderRadius: 6, padding: "8px 12px", color: "#64748b", fontSize: 12, cursor: "pointer", marginBottom: 12, width: "100%" }}>
          + Add alternate mailing address (overrides firm address)
        </button>
      ) : (
        <div style={{ border: "1px solid #e2e8f0", borderRadius: 6, padding: "10px 12px 2px", marginBottom: 12, background: "#f8fafc" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.4 }}>Alternate Mailing Address</span>
            <button type="button" onClick={clearMailing} style={{ background: "none", border: "none", color: "#94a3b8", fontSize: 11, cursor: "pointer" }}>Clear</button>
          </div>
          <Field label="Address Line 1"><input value={form.mailing_line1 || ""} onChange={(e) => set("mailing_line1", e.target.value)} style={inputStyle} /></Field>
          <Field label="Address Line 2"><input value={form.mailing_line2 || ""} onChange={(e) => set("mailing_line2", e.target.value)} style={inputStyle} /></Field>
          <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", gap: 12 }}>
            <Field label="Suburb"><input value={form.mailing_suburb || ""} onChange={(e) => set("mailing_suburb", e.target.value)} style={inputStyle} /></Field>
            <Field label="State">
              <select value={form.mailing_state || ""} onChange={(e) => set("mailing_state", e.target.value)} style={inputStyle}>
                <option value="">—</option>
                {AU_STATES.map((s) => <option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Postcode"><input value={form.mailing_postcode || ""} onChange={(e) => set("mailing_postcode", e.target.value)} maxLength={4} style={inputStyle} /></Field>
          </div>
        </div>
      )}
      <Field label="Tags"><TagInput value={selTags} onChange={setSelTags} allTags={tags} onCreateTag={onCreateTag} /></Field>
      <Field label="Notes"><textarea value={form.notes || ""} onChange={(e) => set("notes", e.target.value)} rows={3} style={{ ...inputStyle, resize: "vertical" }} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button onClick={onClose} style={{ padding: "8px 16px", background: "white", border: "1px solid #d1d5db", borderRadius: 6, cursor: "pointer" }}>Cancel</button>
        <button onClick={handleSave} disabled={saving} style={{ padding: "8px 22px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
          {saving ? "Saving…" : "Save"}
        </button>
      </div>
    </Modal>
  );
}

function ActivityModal({ contacts, initialContactId, onSave, onClose }) {
  const [form, setForm] = useState({ contact_id: initialContactId || "", type: "Phone Call", date: todayStr(), notes: "", follow_up_date: "" });
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const handleSave = async () => {
    if (!form.contact_id) return alert("Please select a contact.");
    setSaving(true); await onSave({ ...form, follow_up_date: form.follow_up_date || null }); setSaving(false);
  };
  return (
    <Modal title="Log Activity" onClose={onClose}>
      <Field label="Contact *">
        <select value={form.contact_id} onChange={(e) => set("contact_id", e.target.value)} style={inputStyle}>
          <option value="">— Select contact —</option>
          {contacts.map((c) => <option key={c.id} value={c.id}>{c.first_name} {c.last_name}{c.firm ? ` (${c.firm.name})` : ""}</option>)}
        </select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Activity Type">
          <select value={form.type} onChange={(e) => set("type", e.target.value)} style={inputStyle}>
            {ACTIVITY_TYPES.map((t) => <option key={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Date"><input type="date" value={form.date} onChange={(e) => set("date", e.target.value)} style={inputStyle} /></Field>
      </div>
      <Field label="Notes"><textarea value={form.notes} onChange={(e) => set("notes", e.target.value)} rows={3} style={{ ...inputStyle, resize: "vertical" }} /></Field>
      <Field label="Set Follow-up Date (optional)"><input type="date" value={form.follow_up_date} onChange={(e) => set("follow_up_date", e.target.value)} style={inputStyle} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button onClick={onClose} style={{ padding: "8px 16px", background: "white", border: "1px solid #d1d5db", borderRadius: 6, cursor: "pointer" }}>Cancel</button>
        <button onClick={handleSave} disabled={saving} style={{ padding: "8px 22px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
          {saving ? "Saving…" : "Log Activity"}
        </button>
      </div>
    </Modal>
  );
}

function FirmModal({ firm, onSave, onClose }) {
  const [form, setForm] = useState(firm || { name: "", type: "Law Firm", address_line1: "", address_line2: "", suburb: "", state: "", postcode: "", phone: "", website: "", notes: "" });
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  return (
    <Modal title={firm ? "Edit Firm" : "Add Firm"} onClose={onClose} width={620}>
      <Field label="Firm Name *"><input value={form.name} onChange={(e) => set("name", e.target.value)} style={inputStyle} /></Field>
      <Field label="Type">
        <select value={form.type} onChange={(e) => set("type", e.target.value)} style={inputStyle}>
          {FIRM_TYPES.map((t) => <option key={t}>{t}</option>)}
        </select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Phone"><input value={form.phone || ""} onChange={(e) => set("phone", e.target.value)} style={inputStyle} /></Field>
        <Field label="Website"><input value={form.website || ""} onChange={(e) => set("website", e.target.value)} style={inputStyle} /></Field>
      </div>
      <div style={{ fontSize: 12, fontWeight: 600, color: "#64748b", textTransform: "uppercase", letterSpacing: 0.4, margin: "8px 0 6px" }}>Mailing Address</div>
      <Field label="Address Line 1"><input value={form.address_line1 || ""} onChange={(e) => set("address_line1", e.target.value)} placeholder="Level 12, 123 Smith Street" style={inputStyle} /></Field>
      <Field label="Address Line 2"><input value={form.address_line2 || ""} onChange={(e) => set("address_line2", e.target.value)} placeholder="Suite / PO Box (optional)" style={inputStyle} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr", gap: 12 }}>
        <Field label="Suburb"><input value={form.suburb || ""} onChange={(e) => set("suburb", e.target.value)} style={inputStyle} /></Field>
        <Field label="State">
          <select value={form.state || ""} onChange={(e) => set("state", e.target.value)} style={inputStyle}>
            <option value="">—</option>
            {AU_STATES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Postcode"><input value={form.postcode || ""} onChange={(e) => set("postcode", e.target.value)} maxLength={4} style={inputStyle} /></Field>
      </div>
      <Field label="Notes"><textarea value={form.notes || ""} onChange={(e) => set("notes", e.target.value)} rows={2} style={{ ...inputStyle, resize: "vertical" }} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button onClick={onClose} style={{ padding: "8px 16px", background: "white", border: "1px solid #d1d5db", borderRadius: 6, cursor: "pointer" }}>Cancel</button>
        <button
          onClick={async () => { if (!form.name) return alert("Name required."); setSaving(true); await onSave(form); setSaving(false); }}
          disabled={saving}
          style={{ padding: "8px 22px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}
        >
          {saving ? "Saving…" : "Save"}
        </button>
      </div>
    </Modal>
  );
}

// ---- Main App ----

export default function BDCRM() {
  const [page, setPage] = useState("dashboard");
  const [selectedContactId, setSelectedContactId] = useState(null);
  const [contacts, setContacts] = useState([]);
  const [firms, setFirms] = useState([]);
  const [tags, setTags] = useState([]);
  const [activities, setActivities] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [contactModal, setContactModal] = useState(null);
  const [activityModal, setActivityModal] = useState(null);
  const [firmModal, setFirmModal] = useState(null);

  const loadAll = useCallback(async () => {
    try {
      setLoading(true);
      const [c, f, t, a] = await Promise.all([
        api("crm_contacts?select=*,crm_firms(id,name,type,address_line1,address_line2,suburb,state,postcode),crm_contact_tags(crm_tags(id,name))&order=last_name.asc,first_name.asc"),
        api("crm_firms?select=*&order=name.asc"),
        api("crm_tags?select=*&order=name.asc"),
        api("crm_activities?select=*&order=date.desc,created_at.desc"),
      ]);
      setContacts(c.map((x) => ({ ...x, firm: x.crm_firms, tags: (x.crm_contact_tags || []).map((ct) => ct.crm_tags).filter(Boolean) })));
      setFirms(f); setTags(t); setActivities(a); setError(null);
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { loadAll(); }, [loadAll]);

  const createTag = async (name) => {
    try {
      const [nt] = await api("crm_tags", { method: "POST", body: JSON.stringify({ name }) });
      setTags((prev) => [...prev, nt].sort((a, b) => a.name.localeCompare(b.name)));
      return nt;
    } catch (e) { return tags.find((t) => t.name.toLowerCase() === name.toLowerCase()); }
  };

  const saveContact = async (data) => {
    const { id, tags: ct, firm, crm_firms, crm_contact_tags, ...rest } = data;
    let cid = id;
    if (id) await api(`crm_contacts?id=eq.${id}`, { method: "PATCH", body: JSON.stringify(rest) });
    else { const [nc] = await api("crm_contacts", { method: "POST", body: JSON.stringify(rest) }); cid = nc.id; }
    if (ct !== undefined) {
      await api(`crm_contact_tags?contact_id=eq.${cid}`, { method: "DELETE" });
      if (ct.length > 0) await api("crm_contact_tags", { method: "POST", body: JSON.stringify(ct.map((t) => ({ contact_id: cid, tag_id: t.id }))) });
    }
    await loadAll(); setContactModal(null);
  };

  const deleteContact = async (id) => {
    if (!confirm("Delete this contact? This cannot be undone.")) return;
    await api(`crm_contacts?id=eq.${id}`, { method: "DELETE" });
    await loadAll(); setPage("contacts"); setSelectedContactId(null);
  };

  const saveFirm = async (data) => {
    const { id, ...rest } = data;
    if (id) await api(`crm_firms?id=eq.${id}`, { method: "PATCH", body: JSON.stringify(rest) });
    else await api("crm_firms", { method: "POST", body: JSON.stringify(rest) });
    await loadAll(); setFirmModal(null);
  };

  const logActivity = async (data) => {
    await api("crm_activities", { method: "POST", body: JSON.stringify(data) });
    await loadAll(); setActivityModal(null);
  };

  const nav = (p, extra = {}) => {
    setPage(p);
    if (extra.contactId !== undefined) setSelectedContactId(extra.contactId);
  };

  const selectedContact = contacts.find((c) => c.id === selectedContactId);
  const contactActivities = activities.filter((a) => a.contact_id === selectedContactId);

  const navItems = [
    { id: "dashboard", label: "Dashboard", icon: "⊞" },
    { id: "contacts",  label: "Contacts",  icon: "👤" },
    { id: "firms",     label: "Firms",     icon: "🏢" },
    { id: "import",    label: "Import",    icon: "📥" },
    { id: "settings",  label: "Settings",  icon: "⚙️" },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", background: "#f8fafc", overflow: "hidden" }}>
      {/* Sidebar */}
      <nav style={{ width: 210, background: "#1e3a5f", color: "white", display: "flex", flexDirection: "column", flexShrink: 0 }}>
        <div style={{ padding: "20px 20px 16px", borderBottom: "1px solid #2d5080" }}>
          <div style={{ fontSize: 17, fontWeight: 700, letterSpacing: "-0.3px" }}>BD CRM</div>
          <div style={{ fontSize: 11, color: "#7fb3d3", marginTop: 2 }}>Forensic Accounting</div>
        </div>
        <div style={{ flex: 1 }}>
          {navItems.map((item) => {
            const active = item.id === "contacts" ? page === "contacts" || page === "contact-detail" : page === item.id;
            return (
              <button
                key={item.id}
                onClick={() => setPage(item.id)}
                style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 20px", background: active ? "#2d5080" : "transparent", color: active ? "white" : "#a8c8e8", border: "none", cursor: "pointer", textAlign: "left", fontSize: 14, width: "100%", transition: "background 0.1s" }}
              >
                <span style={{ fontSize: 16 }}>{item.icon}</span> {item.label}
              </button>
            );
          })}
        </div>
        <div style={{ padding: "12px 16px", borderTop: "1px solid #2d5080" }}>
          <button
            onClick={() => setActivityModal({})}
            style={{ width: "100%", padding: "9px", background: "#2d7dd2", color: "white", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}
          >
            + Log Activity
          </button>
        </div>
      </nav>

      {/* Main Content */}
      <main style={{ flex: 1, overflow: "auto", padding: 24 }}>
        {error && (
          <div style={{ background: "#fee2e2", border: "1px solid #fca5a5", borderRadius: 8, padding: "12px 16px", marginBottom: 16, color: "#991b1b", fontSize: 13 }}>
            ⚠️ <strong>Database error:</strong> {error}
          </div>
        )}
        {loading ? (
          <div style={{ textAlign: "center", padding: 80, color: "#94a3b8", fontSize: 15 }}>Loading…</div>
        ) : (
          <>
            {page === "dashboard" && (
              <Dashboard contacts={contacts} activities={activities} firms={firms} onNav={nav} onLogActivity={() => setActivityModal({})} />
            )}
            {page === "contacts" && (
              <ContactsList
                contacts={contacts} firms={firms} tags={tags}
                onSelect={(c) => nav("contact-detail", { contactId: c.id })}
                onAdd={() => setContactModal("new")}
                onEdit={(c) => setContactModal(c)}
              />
            )}
            {page === "contact-detail" && selectedContact && (
              <ContactDetail
                contact={selectedContact}
                activities={contactActivities}
                firms={firms}
                onEdit={() => setContactModal(selectedContact)}
                onDelete={() => deleteContact(selectedContact.id)}
                onLog={() => setActivityModal({ contactId: selectedContact.id })}
                onBack={() => setPage("contacts")}
              />
            )}
            {page === "firms" && (
              <FirmsList
                firms={firms} contacts={contacts}
                onAdd={() => setFirmModal("new")}
                onEdit={(f) => setFirmModal(f)}
                onSelectContact={(c) => nav("contact-detail", { contactId: c.id })}
              />
            )}
            {page === "import" && (
              <ImportPage firms={firms} tags={tags} onCreateTag={createTag} onImportComplete={loadAll} />
            )}
            {page === "settings" && (
              <SettingsPage tags={tags} onRefresh={loadAll} />
            )}
          </>
        )}
      </main>

      {/* Modals */}
      {contactModal && (
        <ContactModal
          contact={contactModal === "new" ? null : contactModal}
          firms={firms} tags={tags}
          onCreateTag={createTag}
          onSave={saveContact}
          onClose={() => setContactModal(null)}
        />
      )}
      {activityModal && (
        <ActivityModal
          contacts={contacts}
          initialContactId={activityModal.contactId || ""}
          onSave={logActivity}
          onClose={() => setActivityModal(null)}
        />
      )}
      {firmModal && (
        <FirmModal
          firm={firmModal === "new" ? null : firmModal}
          onSave={saveFirm}
          onClose={() => setFirmModal(null)}
        />
      )}
    </div>
  );
}
