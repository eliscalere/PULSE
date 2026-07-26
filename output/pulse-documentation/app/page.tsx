"use client";

import { useEffect, useMemo, useState } from "react";
import { DocumentPage, parsePage } from "./document-reader";

type Document = {
  id: string;
  number: string;
  title: string;
  type: string;
  audience: string;
  pages: number;
  description: string;
  topics: string[];
  textFile: string;
  pdfFile: string;
};

const documents: Document[] = [
  { id: "brand", number: "01", title: "Brand Identity Guide", type: "Brand system", audience: "All PULSE contributors", pages: 9, description: "The visual and verbal system for PULSE, including mark use, color, typography, tone, and applications.", topics: ["Brand", "Color", "Typography"], textFile: "01_PULSE_Brand_Identity_Guide.txt", pdfFile: "01_PULSE_Brand_Identity_Guide.pdf" },
  { id: "brief", number: "02", title: "Product Brief", type: "Product overview", audience: "Leadership and stakeholders", pages: 4, description: "What PULSE is, the operational gap it closes, and how it connects work, status, approvals, and reporting.", topics: ["Product", "SharePoint", "Reporting"], textFile: "02_PULSE_Product_Brief.txt", pdfFile: "02_PULSE_Product_Brief.pdf" },
  { id: "overview", number: "03", title: "General Overview", type: "Orientation", audience: "New users and sponsors", pages: 8, description: "A clear orientation to the PULSE operating model, roles, project-centered flow, and SharePoint context.", topics: ["Orientation", "Roles", "Projects"], textFile: "03_PULSE_General_Overview.txt", pdfFile: "03_PULSE_General_Overview.pdf" },
  { id: "guide", number: "04", title: "Internal User Guide", type: "User reference", audience: "PULSE users", pages: 16, description: "Step-by-step guidance for navigating PULSE and completing common day-to-day work.", topics: ["User guide", "Meetings", "Travel"], textFile: "04_PULSE_Internal_User_Guide.txt", pdfFile: "04_PULSE_Internal_User_Guide.pdf" },
  { id: "technical", number: "05", title: "Technical Reference", type: "Maintainer reference", audience: "Maintainers and administrators", pages: 18, description: "Architecture, SharePoint integration, schema, roles, browser package, validation, and support details.", topics: ["Architecture", "Lists", "Administration"], textFile: "05_PULSE_Technical_Reference.txt", pdfFile: "05_PULSE_Technical_Reference.pdf" },
  { id: "sop", number: "06", title: "Standard Operating Procedures", type: "Controlled procedure", audience: "Users and administrators", pages: 20, description: "Eight operating procedures covering projects, meetings, travel, reviews, tickets, administration, exports, and evidence.", topics: ["SOP", "Document Review", "Tickets"], textFile: "06_PULSE_Standard_Operating_Procedures.txt", pdfFile: "06_PULSE_Standard_Operating_Procedures.pdf" },
  { id: "governance", number: "07", title: "Documentation Verification & Governance", type: "Assurance package", audience: "Documentation owners", pages: 11, description: "Verification standards, controlled-document governance, traceability, release controls, and review criteria.", topics: ["Governance", "Verification", "Traceability"], textFile: "07_PULSE_Documentation_Verification_and_Governance.txt", pdfFile: "07_PULSE_Documentation_Verification_and_Governance.pdf" },
  { id: "packaging", number: "08", title: "Codebase Setup & Firepit Packaging", type: "Technical orientation", audience: "Developers and release owners", pages: 11, description: "How the source is organized, prepared, packaged, validated, and released as a SharePoint/Firepit-ready file.", topics: ["Codebase", "Packaging", "Release"], textFile: "08_PULSE_Codebase_Setup_and_Firepit_Packaging_Guide.txt", pdfFile: "08_PULSE_Codebase_Setup_and_Firepit_Packaging_Guide.pdf" },
];

type Result = { doc: Document; page: number; excerpt: string; score: number };

function normalize(value: string) { return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
function excerptFor(text: string, terms: string[]) {
  const cleaned = text.replace(/\s+/g, " ").trim();
  const at = Math.max(0, ...terms.map((term) => cleaned.toLowerCase().indexOf(term)));
  const start = Math.max(0, at - 150);
  return `${start ? "…" : ""}${cleaned.slice(start, start + 460)}${start + 460 < cleaned.length ? "…" : ""}`;
}

export default function Home() {
  const [active, setActive] = useState("overview");
  const [view, setView] = useState<"docs" | "brand">("docs");
  const [query, setQuery] = useState("");
  const [source, setSource] = useState<Record<string, string>>({});
  const [bootPhase, setBootPhase] = useState<"loading" | "exiting" | "ready">("loading");
  const [isSearching, setIsSearching] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    let cancelled = false;
    let exitTimer: number | undefined;
    let readyTimer: number | undefined;
    const startedAt = performance.now();
    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

    Promise.all(documents.map(async (doc) => [doc.id, await fetch(`/source-text/${doc.textFile}`).then((r) => r.text())] as const))
      .then((loaded) => {
        if (!cancelled) setSource(Object.fromEntries(loaded));
      })
      .catch((error) => {
        console.error("Unable to load the PULSE documentation source library.", error);
      })
      .finally(() => {
        const minimumDisplay = reducedMotion ? 0 : 1250;
        const remaining = Math.max(0, minimumDisplay - (performance.now() - startedAt));
        exitTimer = window.setTimeout(() => {
          if (cancelled) return;
          setBootPhase("exiting");
          readyTimer = window.setTimeout(() => {
            if (!cancelled) setBootPhase("ready");
          }, reducedMotion ? 0 : 240);
        }, remaining);
      });

    return () => {
      cancelled = true;
      if (exitTimer) window.clearTimeout(exitTimer);
      if (readyTimer) window.clearTimeout(readyTimer);
    };
  }, []);

  const selected = documents.find((doc) => doc.id === active) ?? documents[0];
  const selectedPages = useMemo(
    () => (source[selected.id] ?? "").split("\f").filter((page) => page.trim()),
    [selected.id, source],
  );
  const selectedSections = useMemo(
    () => selectedPages.map((page, index) => ({ page: index + 1, title: parsePage(page).title })),
    [selectedPages],
  );
  const currentSection = selectedSections[currentPage - 1];
  const results = useMemo<Result[]>(() => {
    const terms = normalize(query).split(" ").filter((term) => term.length > 1);
    if (!terms.length || !Object.keys(source).length) return [];
    return documents.flatMap((doc) => (source[doc.id] ?? "").split("\f").map((pageText, index) => {
      const body = normalize(pageText); const score = terms.reduce((sum, term) => sum + (body.split(term).length - 1), 0);
      return score ? { doc, page: index + 1, excerpt: excerptFor(pageText, terms), score } : null;
    }).filter((result): result is Result => Boolean(result))).sort((a, b) => b.score - a.score).slice(0, 30);
  }, [query, source]);

  function choose(id: string, page = 1) { setActive(id); setCurrentPage(page); setView("docs"); setQuery(""); setIsSearching(false); setMobileOpen(false); window.scrollTo({ top: 0, behavior: "smooth" }); }
  function openBrand() { setView("brand"); setQuery(""); setIsSearching(false); setMobileOpen(false); window.scrollTo({ top: 0, behavior: "smooth" }); }
  function search(value: string) { setQuery(value); setIsSearching(Boolean(value.trim())); }
  function goToPage(page: number) {
    setCurrentPage(Math.min(Math.max(page, 1), selectedPages.length || 1));
    setMobileOpen(false);
    document.getElementById("reader-top")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  return <>
    {bootPhase !== "ready" && <PulseBootLoader exiting={bootPhase === "exiting"} />}
    <main className={bootPhase === "ready" ? "pulse-app-shell" : "pulse-app-shell pulse-app-shell--booting"} aria-hidden={bootPhase === "ready" ? undefined : true}>
    <aside className={`sidebar ${mobileOpen ? "open" : ""}`}>
      <a className="brand" href="#top" onClick={() => choose("overview")}><img suppressHydrationWarning src="/brand-assets/PULSE_Wordmark_White_Transparent.png" alt="PULSE" /></a>
      <div className="side-label">Documentation</div>
      <nav className="document-nav">{documents.filter((doc) => doc.id !== "brand").map((doc) => {
        const isActive = active === doc.id && !isSearching && view === "docs";
        return <div className="nav-group" key={doc.id}><button className={isActive ? "nav-item active" : "nav-item"} onClick={() => choose(doc.id)}><span>{doc.number}</span><b>{doc.title}</b></button>{isActive && selectedSections.length > 0 && <div className="subnav" aria-label={`${doc.title} sections`}>{selectedSections.map((section) => <button className={currentPage === section.page ? "subnav-item active" : "subnav-item"} aria-current={currentPage === section.page ? "page" : undefined} key={section.page} onClick={() => goToPage(section.page)}><span>{String(section.page).padStart(2, "0")}</span><b>{section.title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "")}</b></button>)}</div>}</div>;
      })}</nav>
      <div className="side-footer"><span className="status-dot" /> Complete source library<br /><small>97 pages · July 2026</small></div>
      <button className={view === "brand" ? "brand-button active" : "brand-button"} onClick={openBrand}><img suppressHydrationWarning src="/brand-assets/PULSE_Dot_Mark_White_Transparent.png" alt="" /><span><b>Brand guidelines</b><small>Identity & assets</small></span><strong>→</strong></button>
    </aside>
    <section className="content" id="top">
      <header className="topbar"><button className="menu" aria-label="Open documentation navigation" onClick={() => setMobileOpen(!mobileOpen)}>☰</button><div className="crumb">PULSE / <span>{isSearching ? "Search" : view === "brand" ? "Brand guidelines" : `${selected.title}${currentSection ? ` / ${currentSection.title}` : ""}`}</span></div>{view === "brand" ? <a suppressHydrationWarning className="source-link" href="/brand-assets/PULSE_Brand_Identity_Guide_v1.3.pdf" download>Download guide PDF ↓</a> : <a suppressHydrationWarning className="source-link" href={`/source-pdfs/${selected.pdfFile}#page=${currentPage}`} target="_blank" rel="noreferrer">Open source page ↗</a>}</header>
      <div className="hero"><div className="eyebrow">PULSE KNOWLEDGE BASE <i /></div><h1>Search PULSE documentation.</h1><p>Ask a plain-language question, find the relevant section, and verify it against the exact source page.</p>
        <label className="search"><span>⌕</span><input value={query} onChange={(event) => search(event.target.value)} placeholder="Ask anything: How do I submit travel? What lists does PULSE use?" /><kbd>⌘ K</kbd></label>
        <div className="quick">Try: <button onClick={() => search("document review")}>document review</button><button onClick={() => search("project tracker")}>project tracker</button><button onClick={() => search("Firepit package")}>Firepit package</button></div>
      </div>
      {isSearching ? <section className="search-view"><div className="section-kicker">SEARCH RESULTS</div><h2>{results.length ? `${results.length} relevant passages` : "No matching passages yet"}</h2><p className="muted">{results.length ? `Results from the complete source text for “${query}”.` : "Try a shorter phrase, an app feature, role, or process name."}</p>
        <div className="results">{results.map((result, index) => <article className="result" key={`${result.doc.id}-${result.page}-${index}`}><div className="result-meta"><span>{result.doc.number}</span>{result.doc.title} · Page {result.page}</div><p>{result.excerpt}</p><div><button onClick={() => choose(result.doc.id, result.page)}>Read this section →</button><a suppressHydrationWarning href={`/source-pdfs/${result.doc.pdfFile}#page=${result.page}`} target="_blank" rel="noreferrer">Open original ↗</a></div></article>)}</div>
      </section> : view === "brand" ? <BrandPage /> : <>
        <section className="document-header"><div><div className="section-kicker">DOCUMENT {selected.number} · {selected.type}</div><h2>{selected.title}</h2><p>{selected.description}</p><div className="meta"><span>{selected.pages} pages</span><span>{selected.audience}</span></div></div><a suppressHydrationWarning className="pdf-card" href={`/source-pdfs/${selected.pdfFile}`} target="_blank" rel="noreferrer"><span>PDF</span><b>View controlled source</b><small>Original portfolio document ↗</small></a></section>
        <section className="reader" id="reader-top"><div className="reader-heading"><div><span className="reading-position">Section {currentPage} of {selectedPages.length || selected.pages}</span><h3>{currentSection?.title ?? "Loading document…"}</h3></div><div className="topics">{selected.topics.map((topic) => <span key={topic}>{topic}</span>)}</div></div>
          {!selectedPages.length ? <div className="loading">Preparing the structured reading view…</div> : <DocumentPage raw={selectedPages[currentPage - 1] ?? selectedPages[0]} pageNumber={currentPage} documentTitle={selected.title} />}
          {selectedPages.length > 1 && <nav className="reader-pagination" aria-label="Document section navigation"><button disabled={currentPage === 1} onClick={() => goToPage(currentPage - 1)}>← Previous section</button><span>{currentPage} / {selectedPages.length}</span><button disabled={currentPage === selectedPages.length} onClick={() => goToPage(currentPage + 1)}>Next section →</button></nav>}
        </section>
      </>}
    </section>
    </main>
  </>;
}

function PulseBootLoader({ exiting }: { exiting: boolean }) {
  return <div className={exiting ? "pulse-doc-loader pulse-doc-loader--exiting" : "pulse-doc-loader"} role="status" aria-live="polite" aria-label="Loading PULSE Documentation">
    <div className="pulse-svgl-wrap">
      <svg className="pulse-svgl" viewBox="0 0 960 360" aria-hidden="true">
        <g className="pulse-svgl-wordmark">
          <text className="pulse-svgl-letter pulse-svgl-letter--p" x="120" y="180">P</text>
          <text className="pulse-svgl-letter pulse-svgl-letter--u" x="300" y="180">U</text>
          <text className="pulse-svgl-letter pulse-svgl-letter--l" x="480" y="180">L</text>
          <text className="pulse-svgl-letter pulse-svgl-letter--s" x="660" y="180">S</text>
          <text className="pulse-svgl-letter pulse-svgl-letter--e" x="840" y="180">E</text>
        </g>
        <circle className="pulse-svgl-dot pulse-svgl-dot--upper" cx="300" cy="76" r="14" />
        <circle className="pulse-svgl-dot pulse-svgl-dot--lower" cx="660" cy="258" r="14" />
      </svg>
    </div>
    <span className="pulse-loader-sr">Loading the PULSE documentation library</span>
  </div>;
}

function BrandPage() {
  const assets = [
    { title: "Primary wordmark — Signal Black", file: "PULSE_Wordmark_Black_Transparent.png", note: "Default dotted wordmark for light backgrounds", dark: false, compact: false },
    { title: "Primary wordmark — Field White", file: "PULSE_Wordmark_White_Transparent.png", note: "Reversed dotted wordmark for dark backgrounds", dark: true, compact: false },
    { title: "Plain wordmark — Signal Black", file: "PULSE_Plain_Wordmark_Black_Transparent.png", note: "Dot-free wordmark for approved lockup use", dark: false, compact: false },
    { title: "Plain wordmark — Field White", file: "PULSE_Plain_Wordmark_White_Transparent.png", note: "Reversed dot-free wordmark", dark: true, compact: false },
    { title: "Secondary lockup — Signal Black", file: "PULSE_Secondary_Lockup_Black_Transparent.png", note: "Two-dot mark plus dot-free wordmark", dark: false, compact: false },
    { title: "Secondary lockup — Field White", file: "PULSE_Secondary_Lockup_White_Transparent.png", note: "Reversed lockup for dark fields", dark: true, compact: false },
    { title: "Dot mark — Signal Black", file: "PULSE_Dot_Mark_Black_Transparent.png", note: "For compact surfaces, icons, and avatars", dark: false, compact: true },
    { title: "Dot mark — Field White", file: "PULSE_Dot_Mark_White_Transparent.png", note: "For compact use on dark surfaces", dark: true, compact: true },
  ];
  return <section className="brand-page">
    <div className="brand-intro"><div><div className="section-kicker">PULSE BRAND IDENTITY · V1.3</div><h2>The operating rhythm</h2><p>PULSE turns scattered activity into a current, shared, and actionable view of the work. The identity is deliberately restrained: black and white carry the brand, while interface color clarifies action and state.</p></div><div className="brand-downloads"><a suppressHydrationWarning className="download-guide" href="/brand-assets/PULSE_Brand_Assets_v1.3.zip" download>Download all assets <span>↓</span></a><a suppressHydrationWarning href="/brand-assets/PULSE_Brand_Identity_Guide_v1.3.pdf" download>Brand guide PDF</a><a suppressHydrationWarning href="/brand-assets/PULSE_Brand_Assets_README.txt" download>Usage README</a></div></div>
    <div className="brand-principles"><article><span>01</span><h3>Signal in. Action out.</h3><p>The upper dot represents an incoming update or observation. The lower dot completes the pulse: a decision, handoff, approval, or action.</p></article><article><span>02</span><h3>Calm, direct, operational.</h3><p>Use familiar words. State the action, owner, and timing. Be specific without adding drama.</p></article><article><span>03</span><h3>Protect the signal.</h3><p>Keep the mark proportional, clear of nearby elements, and free from effects, status colors, or repositioned dots.</p></article></div>
    <div className="brand-section-title"><div><div className="section-kicker">COLOR</div><h3>Monochrome identity, purposeful interface color</h3></div><p>Use blue for links and current selection; reserve green, amber, and red for status.</p></div>
    <div className="palette"><div className="swatch black"><b>Signal Black</b><span>#070708</span></div><div className="swatch white"><b>Field White</b><span>#FFFFFF</span></div><div className="swatch blue"><b>Command Blue</b><span>#2F66FF</span></div><div className="swatch green"><b>Status Green</b><span>#2B7A5E</span></div><div className="swatch amber"><b>Warning Amber</b><span>#C77800</span></div><div className="swatch red"><b>Critical Red</b><span>#B42318</span></div></div>
    <div className="brand-section-title asset-heading"><div><div className="section-kicker">DOWNLOADS</div><h3>Identity assets</h3></div><p>Use the supplied files as-is. The full guide contains construction, clear-space, and minimum-size rules.</p></div>
    <div className="asset-grid">{assets.map((asset) => <article className={`asset-card ${asset.dark ? "dark" : ""} ${asset.compact ? "compact-asset" : ""}`} key={asset.file}><div className="asset-preview"><img suppressHydrationWarning src={`/brand-assets/${asset.file}`} alt={asset.title} /></div><div><b>{asset.title}</b><p>{asset.note}</p><a suppressHydrationWarning href={`/brand-assets/${asset.file}`} download>Download PNG ↓</a></div></article>)}</div>
  </section>;
}
