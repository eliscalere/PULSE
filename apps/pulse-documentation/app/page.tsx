"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
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
  { id: "focused", number: "09", title: "Focused Tools & Package Delivery", type: "Technical orientation", audience: "Site owners and release owners", pages: 9, description: "The eight delivered packages, what My Travel, the travel calendar, Tickets, and PULSE CODE are for, how a focused tool resolves its SharePoint site, and how to publish and verify one.", topics: ["Packages", "Travel", "Hosting"], textFile: "09_PULSE_Focused_Tools_and_Package_Delivery.txt", pdfFile: "09_PULSE_Focused_Tools_and_Package_Delivery.pdf" },
  { id: "devscale", number: "10", title: "Development Environment & Scale Validation", type: "Follow-on work", audience: "Developers and maintainers", pages: 9, description: "PULSE has no development environment. How to stand up a disposable development site, load it past the paging and view-threshold ceilings already in the code, measure what breaks, and evaluate Dataverse as an alternative store.", topics: ["Development", "Scale", "Dataverse"], textFile: "10_PULSE_Development_Environment_and_Scale_Validation.txt", pdfFile: "10_PULSE_Development_Environment_and_Scale_Validation.pdf" },
];

/* Interface reference. Captures are produced by
   apps/PULSE/scripts/capture-documentation-screenshots.cjs against the current
   build, using synthetic records and neutral role display names. Update this
   date whenever the capture set is regenerated. */
const SCREEN_CAPTURE_DATE = "July 2026";

type ScreenGroup = { group: string; blurb: string; screens: { file: string; title: string; route: string; note: string }[] };

const screenGroups: ScreenGroup[] = [
  {
    group: "Situational awareness",
    blurb: "Where a session starts: personal work, then the portfolio-wide picture.",
    screens: [
      { file: "01-dashboard.png", title: "Dashboard", route: "#/dashboard", note: "Counts, quick access into each area, and assigned work across projects." },
      { file: "02-overview-portfolio.png", title: "Overview — portfolio", route: "#/overview", note: "Team and portfolio roll-up for status, workload, and approvals." },
    ],
  },
  {
    group: "Project workspaces",
    blurb: "The project is the front door: people, files, work items, and reporting all sit in project context.",
    screens: [
      { file: "03-projects-workspaces.png", title: "All workspaces", route: "#/projects", note: "Project list with health and entry points into each workspace." },
      { file: "04-project-workspace.png", title: "Workspace home", route: "#/projects/<code>", note: "Project landing view with the workspace section rail." },
      { file: "05-project-tracker.png", title: "Tracker", route: "#/projects/<code>/tracker", note: "Tasks and milestones with owner, dates, health, and timeline and risk tabs." },
      { file: "06-project-checklist.png", title: "Checklist", route: "#/projects/<code>/checklist", note: "Column-based checklist for repeatable delivery steps." },
      { file: "07-project-documents.png", title: "Documents", route: "#/projects/<code>/documents", note: "Project document surface backed by the SharePoint library." },
      { file: "08-project-people.png", title: "People", route: "#/projects/<code>/people", note: "Assigned members and roles for the workspace." },
    ],
  },
  {
    group: "Recurring operations",
    blurb: "The weekly rhythm and the formal review path.",
    screens: [
      { file: "09-weekly-meeting.png", title: "Weekly meeting", route: "#/weekly", note: "Live session with attendance, minutes, project updates, and around-the-room." },
      { file: "14-document-review.png", title: "Document review", route: "#/docreview", note: "Review packages moving through concurrence and signature." },
    ],
  },
  {
    group: "Travel",
    blurb: "Request, track, and close out travel, including the standalone Firepit tools.",
    screens: [
      { file: "10-travel-request-form.png", title: "Travel request form", route: "#/travel/submit", note: "Guided TDY, conference, training, and leave request intake." },
      { file: "11-travel-my-travel.png", title: "My travel", route: "#/travel/mytravel", note: "A traveler's own requests by state, with concurrence and charge-object status." },
      { file: "12-travel-calendar.png", title: "Travel calendar", route: "#/travel/calendar", note: "Team travel and events on a shared calendar." },
      { file: "13-travel-debrief.png", title: "Travel debrief", route: "#/travel/debrief", note: "Post-trip debrief capture against an approved request." },
    ],
  },
  {
    group: "Support and administration",
    blurb: "Intake, preferences, configuration, and the activity record.",
    screens: [
      { file: "15-tickets.png", title: "Tickets", route: "#/tickets", note: "Blockers, bugs, access needs, and questions tracked to resolution." },
      { file: "16-admin.png", title: "Admin", route: "#/admin", note: "SharePoint setup, users and roles, configuration, and diagnostics." },
      { file: "17-notification-settings.png", title: "Notification settings", route: "#/notification-settings", note: "Per-user areas, tone, and channel preferences." },
      { file: "18-logs.png", title: "Activity log", route: "#/logs", note: "Audit record of actions by actor, area, and time." },
    ],
  },
];

type Result = { doc: Document; page: number; sectionTitle: string; excerpt: string; score: number; hits: number };

/* Splits an excerpt around matched terms so the reader can see why a passage
   was returned without re-reading it. Longest terms first, so "travel request"
   highlights as one span rather than fragmenting on "travel". */
function highlight(text: string, terms: string[]) {
  if (!terms.length) return [text];
  const pattern = new RegExp(`(${terms.slice().sort((a, b) => b.length - a.length).map((term) => term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})`, "gi");
  return text.split(pattern).filter((part) => part !== "");
}

function normalize(value: string) { return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }
function excerptFor(text: string, terms: string[]) {
  const cleaned = text.replace(/\s+/g, " ").trim();
  const at = Math.max(0, ...terms.map((term) => cleaned.toLowerCase().indexOf(term)));
  const start = Math.max(0, at - 150);
  return `${start ? "…" : ""}${cleaned.slice(start, start + 460)}${start + 460 < cleaned.length ? "…" : ""}`;
}

function getAssetUrl(url: string): string {
  if (typeof window !== "undefined") {
    if ((window as any).getAssetUrl) {
      return (window as any).getAssetUrl(url);
    }
    if ((window as any).__PULSE_ASSETS__?.[url]) {
      return (window as any).__PULSE_ASSETS__[url];
    }
  }
  return url;
}

export default function Home() {
  const [active, setActive] = useState("overview");
  const [view, setView] = useState<"docs" | "brand" | "screens" | "flows">("docs");
  const [query, setQuery] = useState("");
  const [source, setSource] = useState<Record<string, string>>({});
  const [bootPhase, setBootPhase] = useState<"loading" | "exiting" | "ready">("loading");
  const [isSearching, setIsSearching] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  /* The whole document renders in one scroll. activeSection is what the reader
     is looking at, tracked by observer rather than set by clicking, so the
     sidebar stays truthful when the user scrolls instead of navigating. */
  const [activeSection, setActiveSection] = useState(1);
  const [pendingSection, setPendingSection] = useState<number | null>(null);
  const searchInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    let readyTimer: number | undefined;

    const finish = () => {
      if (cancelled) return;
      setBootPhase((phase) => (phase === "loading" ? "exiting" : phase));
      readyTimer = window.setTimeout(() => {
        if (!cancelled) setBootPhase("ready");
      }, 50);
    };

    /* The packaged single-file build carries every source document inline, so
       read it straight off the global instead of round-tripping through fetch.
       Only the dev server needs the network path. */
    const inlined = (window as unknown as { __PULSE_SOURCE_TEXT__?: Record<string, string> }).__PULSE_SOURCE_TEXT__;
    if (inlined) {
      const loaded = documents
        .map((doc) => [doc.id, inlined[`/source-text/${doc.textFile}`]] as const)
        .filter((entry): entry is readonly [string, string] => typeof entry[1] === "string");
      if (loaded.length) {
        setSource(Object.fromEntries(loaded));
        finish();
        return () => {
          cancelled = true;
          if (readyTimer) window.clearTimeout(readyTimer);
        };
      }
    }

    /* Never let a stalled or failed source load strand the reader behind the
       boot overlay — reveal the shell regardless and let content fill in. */
    const safety = window.setTimeout(finish, 4000);

    Promise.all(documents.map(async (doc) => [doc.id, await fetch(`/source-text/${doc.textFile}`).then((r) => r.text())] as const))
      .then((loaded) => {
        if (!cancelled) setSource(Object.fromEntries(loaded));
      })
      .catch((error) => {
        console.error("Unable to load the PULSE documentation source library.", error);
      })
      .finally(() => {
        window.clearTimeout(safety);
        finish();
      });

    return () => {
      cancelled = true;
      window.clearTimeout(safety);
      if (readyTimer) window.clearTimeout(readyTimer);
    };
  }, []);

  const scrollToSection = useCallback((page: number, behavior: ScrollBehavior = "smooth") => {
    const target = document.getElementById(`section-${page}`);
    if (!target) return false;
    target.scrollIntoView({ behavior, block: "start" });
    return true;
  }, []);

  /* Deep link from a search result: the target document's sections are not in
     the DOM on the click, so the scroll waits for the render that follows. */
  useEffect(() => {
    if (pendingSection === null) return;
    let frame = 0;
    let attempts = 0;
    const tryScroll = () => {
      if (scrollToSection(pendingSection, "auto")) {
        setActiveSection(pendingSection);
        setPendingSection(null);
        return;
      }
      if (attempts++ < 30) frame = window.requestAnimationFrame(tryScroll);
      else setPendingSection(null);
    };
    frame = window.requestAnimationFrame(tryScroll);
    return () => window.cancelAnimationFrame(frame);
  }, [pendingSection, scrollToSection]);

  /* Cmd/Ctrl-K focuses search, matching the hint already shown in the field. */
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
      }
      if (event.key === "Escape" && document.activeElement === searchInputRef.current) {
        searchInputRef.current?.blur();
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
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
  const currentSection = selectedSections[activeSection - 1];

  /* Scroll-spy. The current section is the last one whose top has passed under
     the sticky header — computed from rects rather than intersection ratios,
     because several sections are visible at once on a tall viewport and
     "topmost intersecting" then lags a section behind the heading on screen.
     Reads are batched into a frame so scrolling stays cheap. */
  useEffect(() => {
    if (view !== "docs" || isSearching || !selectedPages.length) return;
    const nodes = Array.from(document.querySelectorAll<HTMLElement>("[data-section]"));
    if (!nodes.length) return;

    let frame = 0;
    const compute = () => {
      frame = 0;
      const line = 108; // sticky topbar plus a little breathing room
      let current = Number(nodes[0].dataset.section) || 1;
      for (const node of nodes) {
        if (node.getBoundingClientRect().top - line <= 0) current = Number(node.dataset.section);
        else break;
      }
      /* At the very bottom the last section is current even if its top never
         crossed the line, which happens when the final section is short. */
      if (window.innerHeight + window.scrollY >= document.body.scrollHeight - 4) {
        current = Number(nodes[nodes.length - 1].dataset.section);
      }
      setActiveSection(current);
    };
    const onScroll = () => { if (!frame) frame = window.requestAnimationFrame(compute); };

    compute();
    window.addEventListener("scroll", onScroll, { passive: true });
    window.addEventListener("resize", onScroll);
    return () => {
      if (frame) window.cancelAnimationFrame(frame);
      window.removeEventListener("scroll", onScroll);
      window.removeEventListener("resize", onScroll);
    };
  }, [view, isSearching, selectedPages, selected.id]);
  const results = useMemo<Result[]>(() => {
    const terms = normalize(query).split(" ").filter((term) => term.length > 1);
    if (!terms.length || !Object.keys(source).length) return [];
    return documents.flatMap((doc) => (source[doc.id] ?? "").split("\f").filter((page) => page.trim()).map((pageText, index) => {
      const body = normalize(pageText);
      const hits = terms.reduce((sum, term) => sum + (body.split(term).length - 1), 0);
      if (!hits) return null;
      const sectionTitle = parsePage(pageText).title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "");
      /* A term in the section title means the passage is about the query, not
         merely mentions it, so weight those above raw frequency. */
      const titleBonus = terms.filter((term) => normalize(sectionTitle).includes(term)).length * 6;
      return { doc, page: index + 1, sectionTitle, excerpt: excerptFor(pageText, terms), score: hits + titleBonus, hits };
    }).filter((result): result is Result => Boolean(result))).sort((a, b) => b.score - a.score).slice(0, 30);
  }, [query, source]);

  const searchTerms = useMemo(() => normalize(query).split(" ").filter((term) => term.length > 1), [query]);
  const resultsByDoc = useMemo(() => {
    const groups = new Map<string, { doc: Document; items: Result[] }>();
    for (const result of results) {
      const group = groups.get(result.doc.id) ?? { doc: result.doc, items: [] };
      group.items.push(result);
      groups.set(result.doc.id, group);
    }
    return Array.from(groups.values());
  }, [results]);

  function choose(id: string, page = 1) {
    const sameDoc = id === active && view === "docs" && !isSearching;
    setActive(id); setView("docs"); setQuery(""); setIsSearching(false); setMobileOpen(false);
    if (page > 1) {
      /* Same document already on screen: scroll now. Different document: defer
         until its sections exist. */
      if (sameDoc && scrollToSection(page)) setActiveSection(page);
      else setPendingSection(page);
      return;
    }
    setActiveSection(1);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
  function openBrand() { setView("brand"); setQuery(""); setIsSearching(false); setMobileOpen(false); window.scrollTo({ top: 0, behavior: "smooth" }); }
  function openScreens() { setView("screens"); setQuery(""); setIsSearching(false); setMobileOpen(false); window.scrollTo({ top: 0, behavior: "smooth" }); }
  function openFlows() { setView("flows"); setQuery(""); setIsSearching(false); setMobileOpen(false); window.scrollTo({ top: 0, behavior: "smooth" }); }
  function search(value: string) { setQuery(value); setIsSearching(Boolean(value.trim())); }
  function goToSection(page: number) {
    setMobileOpen(false);
    setActiveSection(page);
    scrollToSection(page);
  }

  return <>
    {bootPhase !== "ready" && <PulseBootLoader exiting={bootPhase === "exiting"} />}
    <main className={bootPhase === "ready" ? "pulse-app-shell" : "pulse-app-shell pulse-app-shell--booting"} aria-hidden={bootPhase === "ready" ? undefined : true}>
    <aside className={`sidebar ${mobileOpen ? "open" : ""}`}>
      <a className="brand" href="#top" onClick={() => choose("overview")}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Wordmark_White_Transparent.png")} alt="PULSE" /></a>
      <div className="side-label">Documentation</div>
      <nav className="document-nav">{documents.filter((doc) => doc.id !== "brand").map((doc) => {
        const isActive = active === doc.id && !isSearching && view === "docs";
        return <div className="nav-group" key={doc.id}><button className={isActive ? "nav-item active" : "nav-item"} onClick={() => choose(doc.id)}><span>{doc.number}</span><b>{doc.title}</b></button>{isActive && selectedSections.length > 0 && <div className="subnav" aria-label={`${doc.title} sections`}>{selectedSections.map((section) => <button className={activeSection === section.page ? "subnav-item active" : "subnav-item"} aria-current={activeSection === section.page ? "true" : undefined} key={section.page} onClick={() => goToSection(section.page)}><span>{String(section.page).padStart(2, "0")}</span><b>{section.title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "")}</b></button>)}</div>}</div>;
      })}</nav>
      <div className="side-footer"><span className="status-dot" /> Complete source library<br /><small>{documents.reduce((sum, doc) => sum + doc.pages, 0)} pages · July 2026</small></div>
      <button className={view === "flows" ? "brand-button active" : "brand-button"} onClick={openFlows}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Dot_Mark_White_Transparent.png")} alt="" /><span><b>Process flows</b><small>How work moves</small></span><strong>→</strong></button>
      <button className={view === "screens" ? "brand-button active" : "brand-button"} onClick={openScreens}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Dot_Mark_White_Transparent.png")} alt="" /><span><b>Interface reference</b><small>Current screens</small></span><strong>→</strong></button>
      <button className={view === "brand" ? "brand-button active" : "brand-button"} onClick={openBrand}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Dot_Mark_White_Transparent.png")} alt="" /><span><b>Brand guidelines</b><small>Identity & assets</small></span><strong>→</strong></button>
    </aside>
    <section className="content" id="top">
      <header className="topbar"><button className="menu" aria-label="Open documentation navigation" onClick={() => setMobileOpen(!mobileOpen)}>☰</button><div className="crumb">PULSE / <span>{isSearching ? "Search" : view === "brand" ? "Brand guidelines" : view === "screens" ? "Interface reference" : view === "flows" ? "Process flows" : `${selected.title}${currentSection ? ` / ${currentSection.title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "")}` : ""}`}</span></div>{view === "brand" ? <a suppressHydrationWarning className="source-link" href={getAssetUrl("/brand-assets/PULSE_Brand_Identity_Guide_v1.3.pdf")} download>Download guide PDF ↓</a> : view === "screens" ? <span className="source-link source-link--static">Captured {SCREEN_CAPTURE_DATE}</span> : view === "flows" ? <span className="source-link source-link--static">Derived from current source</span> : <a suppressHydrationWarning className="source-link" href={getAssetUrl(`/source-pdfs/${selected.pdfFile}`)} target="_blank" rel="noreferrer">Open source page ↗</a>}</header>
      <div className="hero"><div className="eyebrow">PULSE KNOWLEDGE BASE <i /></div><h1>Search PULSE documentation.</h1><p>Ask a plain-language question, find the relevant section, and verify it against the exact source page.</p>
        <label className="search"><span>⌕</span><input ref={searchInputRef} value={query} onChange={(event) => search(event.target.value)} placeholder="Ask anything: How do I submit travel? What lists does PULSE use?" aria-label="Search the PULSE documentation source library" />{query ? <button type="button" className="search-clear" onClick={() => search("")} aria-label="Clear search">✕</button> : <kbd>⌘ K</kbd>}</label>
        <div className="quick">Try: <button onClick={() => search("document review")}>document review</button><button onClick={() => search("project tracker")}>project tracker</button><button onClick={() => search("Firepit package")}>Firepit package</button></div>
      </div>
      {isSearching ? <section className="search-view"><div className="section-kicker">SEARCH RESULTS</div><h2>{results.length ? `${results.length} relevant passages` : "No matching passages yet"}</h2><p className="muted">{results.length ? `Results from the complete source text for “${query}”.` : "Try a shorter phrase, an app feature, role, or process name."}</p>
        {!results.length && <div className="search-empty"><b>Nothing matched “{query}”.</b><p>Search covers the full text of every document. Try a feature name (<button onClick={() => search("travel debrief")}>travel debrief</button>), a SharePoint list (<button onClick={() => search("PULSE Issues")}>PULSE Issues</button>), or a process (<button onClick={() => search("site resolution")}>site resolution</button>).</p></div>}
        <div className="results-grouped">{resultsByDoc.map((group) => <div className="result-group" key={group.doc.id}>
          <div className="result-group-head"><span className="result-group-num">{group.doc.number}</span><b>{group.doc.title}</b><span className="result-group-count">{group.items.length} {group.items.length === 1 ? "passage" : "passages"}</span></div>
          {group.items.map((result, index) => <article className="result" key={`${result.doc.id}-${result.page}-${index}`}>
            <button className="result-open" onClick={() => choose(result.doc.id, result.page)}>
              <div className="result-head"><span className="result-section">Section {String(result.page).padStart(2, "0")}</span><b>{result.sectionTitle}</b><span className="result-hits">{result.hits} {result.hits === 1 ? "match" : "matches"}</span></div>
              <p>{highlight(result.excerpt, searchTerms).map((part, partIndex) => searchTerms.some((term) => part.toLowerCase() === term.toLowerCase()) ? <mark key={partIndex}>{part}</mark> : <span key={partIndex}>{part}</span>)}</p>
            </button>
            <a suppressHydrationWarning className="result-source" href={getAssetUrl(`/source-pdfs/${result.doc.pdfFile}`)} target="_blank" rel="noreferrer">Open controlled source ↗</a>
          </article>)}
        </div>)}</div>
      </section> : view === "brand" ? <BrandPage /> : view === "screens" ? <ScreensPage /> : view === "flows" ? <FlowsPage /> : <>
        <section className="document-header"><div><div className="section-kicker">DOCUMENT {selected.number} · {selected.type}</div><h2>{selected.title}</h2><p>{selected.description}</p><div className="meta"><span>{selected.pages} pages</span><span>{selected.audience}</span></div></div><a suppressHydrationWarning className="pdf-card" href={getAssetUrl(`/source-pdfs/${selected.pdfFile}`)} target="_blank" rel="noreferrer"><span>PDF</span><b>View controlled source</b><small>Original portfolio document ↗</small></a></section>
        <section className="reader" id="reader-top">
          <div className="reader-heading"><div><span className="reading-position">{selectedPages.length || selected.pages} sections · continuous</span><h3>Read straight through, or jump from the sidebar</h3></div><div className="topics">{selected.topics.map((topic) => <span key={topic}>{topic}</span>)}</div></div>
          {!selectedPages.length
            ? <div className="reader-skeleton" aria-label="Preparing the structured reading view"><span /><span /><span /><span /></div>
            : <div className="reader-stream">{selectedPages.map((raw, index) => <div className="reader-section" data-section={index + 1} id={`section-${index + 1}`} key={`${selected.id}-${index}`}><DocumentPage raw={raw} pageNumber={index + 1} documentTitle={selected.title} /></div>)}</div>}
          {selectedPages.length > 1 && <div className="reader-end"><span>End of {selected.title}</span><button onClick={() => { setActiveSection(1); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Back to top ↑</button></div>}
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

/* ---------- process flows ----------
   Each flow is a left-to-right sequence of states with optional exception
   branches. States and labels are taken from the values the application
   actually displays, so a diagram cannot drift from the product by being
   redrawn prettier than it is. */

type FlowNode = { id: string; label: string; note?: string; tone?: "start" | "step" | "done" | "exception" };
type FlowEdge = { from: string; to: string; label?: string; kind?: "loop" | "exception" };
/* layout: "row" suits short lifecycles. "column" suits ordered fallback chains,
   where every step carries a condition label — laid out horizontally those
   labels have to fit a 46px gap and end up overlapping the boxes. */
type Flow = { id: string; title: string; blurb: string; source: string; layout?: "row" | "column"; lanes: FlowNode[][]; edges: FlowEdge[] };

const flows: Flow[] = [
  {
    id: "travel",
    title: "Travel request",
    blurb: "A request moves through approval and funding before travel, and is not closed until every traveller has filed a debrief.",
    source: "Travel request status values and the travel module",
    lanes: [
      [{ id: "t1", label: "New request", note: "Travel → Submit", tone: "start" }],
      [{ id: "t2", label: "Pending", note: "Awaiting approval" }],
      [{ id: "t3", label: "Pending Finance", note: "Funding and charge object" }],
      [{ id: "t4", label: "Approved", note: "Appears on the calendar" }],
      [{ id: "t5", label: "Debrief filed", note: "One per traveller", tone: "done" }],
    ],
    edges: [
      { from: "t1", to: "t2" }, { from: "t2", to: "t3" }, { from: "t3", to: "t4" }, { from: "t4", to: "t5" },
      { from: "t2", to: "x1", label: "Denied · Cancelled · Withdrawn", kind: "exception" },
      { from: "t4", to: "x2", label: "Revoked", kind: "exception" },
    ],
  },
  {
    id: "docreview",
    title: "Document review",
    blurb: "Formal review is a controlled loop: requested changes produce a new revision and the review runs again. Storing a file under a project is not review.",
    source: "Document Review workflow values",
    lanes: [
      [{ id: "d1", label: "Not Started", tone: "start" }],
      [{ id: "d2", label: "In Review", note: "Reviewers decide" }],
      [{ id: "d3", label: "Review Complete" }],
      [{ id: "d4", label: "Awaiting Final Pack", note: "Then Signing in Progress" }],
      [{ id: "d5", label: "Signed", note: "Then Archived", tone: "done" }],
    ],
    edges: [
      { from: "d1", to: "d2" }, { from: "d2", to: "d3" }, { from: "d3", to: "d4" }, { from: "d4", to: "d5" },
      { from: "d2", to: "d2loop", label: "Changes Requested → new revision", kind: "loop" },
    ],
  },
  {
    id: "ticket",
    title: "Support ticket",
    blurb: "Tickets and issue reports are the same records in the PULSE Issues list, so a ticket raised in the standalone tool is the one the full application shows.",
    source: "Ticket status values and the repository mapping",
    lanes: [
      [{ id: "k1", label: "Open", note: "Raised by any user", tone: "start" }],
      [{ id: "k2", label: "In Progress", note: "Being worked" }],
      [{ id: "k3", label: "Resolved", note: "Outcome recorded", tone: "done" }],
    ],
    edges: [{ from: "k1", to: "k2" }, { from: "k2", to: "k3" }],
  },
  {
    id: "site",
    title: "Site resolution in a web part",
    blurb: "A focused tool in a SharePoint iframe cannot see the page context directly. It tries six sources in order, and local-only mode is the deliberate signal that none succeeded.",
    source: "getSiteUrl in the SharePoint adapter",
    layout: "column",
    lanes: [
      [{ id: "s1", label: "Own page context", note: "Direct SharePoint page", tone: "start" }],
      [{ id: "s2", label: "Configured site", note: "manualSharePointSiteUrl" }],
      [{ id: "s3", label: "Parent window", note: "Normal web-part path" }],
      [{ id: "s4", label: "Top window", note: "Nested frame" }],
      [{ id: "s5", label: "Cached address", note: "From a previous boot" }],
      [{ id: "s6", label: "Local only", note: "Not a system of record", tone: "exception" }],
    ],
    edges: [
      { from: "s1", to: "s2", label: "not found" }, { from: "s2", to: "s3", label: "not set" },
      { from: "s3", to: "s4", label: "not found" }, { from: "s4", to: "s5", label: "not found" },
      { from: "s5", to: "s6", label: "not cached" },
    ],
  },
];

const FLOW_GEOMETRY = { boxW: 176, boxH: 74, gapX: 46, top: 34, exceptionY: 152 };

const COLUMN_GEOMETRY = { boxW: 296, boxH: 64, gapY: 38, labelGap: 18 };

/* Vertical variant: each step stacked, with the condition that causes the
   fall-through to the next step set beside the connector. */
function FlowColumn({ flow }: { flow: Flow }) {
  const { boxW, boxH, gapY, labelGap } = COLUMN_GEOMETRY;
  const width = boxW + 46; // just enough for the step index beside each box
  const height = flow.lanes.length * (boxH + gapY) - gapY;
  const yFor = (index: number) => index * (boxH + gapY);
  const rail = 34;

  return <svg viewBox={`-2 -2 ${width + 4} ${height + 4}`} role="img" aria-label={`${flow.title}: ${flow.lanes.map((lane) => lane[0].label).join(", then ")}`}>
    <defs>
      <marker id={`carrow-${flow.id}`} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
        <path d="M0,1 L9,5 L0,9 z" fill="var(--graphite)" />
      </marker>
    </defs>

    {flow.lanes.slice(0, -1).map((lane, index) => {
      const y1 = yFor(index) + boxH;
      const y2 = yFor(index + 1);
      const edge = flow.edges.find((candidate) => candidate.from === lane[0].id && candidate.to === flow.lanes[index + 1][0].id);
      return <g key={`clink-${index}`}>
        <line x1={rail} y1={y1} x2={rail} y2={y2 - 3} stroke="var(--graphite)" strokeWidth="1.5" markerEnd={`url(#carrow-${flow.id})`} />
        {edge?.label && <text x={rail + labelGap} y={(y1 + y2) / 2 + 3} className="flow-edge-label">{edge.label}</text>}
      </g>;
    })}

    {flow.lanes.map((lane, index) => {
      const node = lane[0];
      const y = yFor(index);
      return <g key={node.id} className={`flow-node flow-node--${node.tone ?? "step"}`}>
        <rect x="0" y={y} width={boxW} height={boxH} rx="2" />
        <text x="15" y={y + (node.note ? 26 : 38)} className="flow-node-label">{node.label}</text>
        {node.note && <text x="15" y={y + 45} className="flow-node-note">{node.note}</text>}
        <text x={boxW + 16} y={y + (node.note ? 33 : 38)} className="flow-step-index">{String(index + 1).padStart(2, "0")}</text>
      </g>;
    })}
  </svg>;
}

function FlowDiagram({ flow }: { flow: Flow }) {
  if (flow.layout === "column") {
    return <figure className="flow">
      <figcaption><b>{flow.title}</b><p>{flow.blurb}</p></figcaption>
      <div className="flow-canvas flow-canvas--column"><FlowColumn flow={flow} /></div>
      <div className="flow-source">Source: {flow.source}</div>
    </figure>;
  }
  const columns = flow.lanes.length;
  const width = columns * FLOW_GEOMETRY.boxW + (columns - 1) * FLOW_GEOMETRY.gapX;
  const hasException = flow.edges.some((edge) => edge.kind === "exception");
  const hasLoop = flow.edges.some((edge) => edge.kind === "loop");
  const height = FLOW_GEOMETRY.top + FLOW_GEOMETRY.boxH + (hasException ? 96 : hasLoop ? 78 : 26);
  const xFor = (index: number) => index * (FLOW_GEOMETRY.boxW + FLOW_GEOMETRY.gapX);
  const midY = FLOW_GEOMETRY.top + FLOW_GEOMETRY.boxH / 2;

  return <figure className="flow">
    <figcaption><b>{flow.title}</b><p>{flow.blurb}</p></figcaption>
    <div className="flow-canvas">
      {/* 2px of viewBox padding keeps the outermost 1px strokes from being
          clipped at the edges once the SVG scales to its container. */}
      <svg viewBox={`-2 0 ${width + 4} ${height}`} role="img" aria-label={`${flow.title}: ${flow.lanes.map((lane) => lane[0].label).join(", then ")}`}>
        <defs>
          <marker id={`arrow-${flow.id}`} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M0,1 L9,5 L0,9 z" fill="var(--graphite)" />
          </marker>
          <marker id={`arrow-x-${flow.id}`} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
            <path d="M0,1 L9,5 L0,9 z" fill="#9aa0a8" />
          </marker>
        </defs>

        {flow.lanes.slice(0, -1).map((_, index) => {
          const x1 = xFor(index) + FLOW_GEOMETRY.boxW;
          const x2 = xFor(index + 1);
          const edge = flow.edges.find((candidate) => candidate.from === flow.lanes[index][0].id && candidate.to === flow.lanes[index + 1][0].id);
          return <g key={`link-${index}`}>
            <line x1={x1} y1={midY} x2={x2 - 3} y2={midY} stroke="var(--graphite)" strokeWidth="1.5" markerEnd={`url(#arrow-${flow.id})`} />
            {edge?.label && <text x={(x1 + x2) / 2} y={midY - 9} textAnchor="middle" className="flow-edge-label">{edge.label}</text>}
          </g>;
        })}

        {flow.lanes.map((lane, index) => {
          const node = lane[0];
          const x = xFor(index);
          return <g key={node.id} className={`flow-node flow-node--${node.tone ?? "step"}`}>
            <rect x={x} y={FLOW_GEOMETRY.top} width={FLOW_GEOMETRY.boxW} height={FLOW_GEOMETRY.boxH} rx="2" />
            <text x={x + 14} y={FLOW_GEOMETRY.top + (node.note ? 30 : 42)} className="flow-node-label">{node.label}</text>
            {node.note && <text x={x + 14} y={FLOW_GEOMETRY.top + 51} className="flow-node-note">{node.note}</text>}
          </g>;
        })}

        {hasLoop && (() => {
          const loop = flow.edges.find((edge) => edge.kind === "loop")!;
          const index = flow.lanes.findIndex((lane) => lane[0].id === loop.from);
          const x = xFor(index);
          const y = FLOW_GEOMETRY.top + FLOW_GEOMETRY.boxH;
          return <g>
            <path d={`M ${x + 30} ${y} V ${y + 34} H ${x + FLOW_GEOMETRY.boxW - 20} V ${y + 3}`} fill="none" stroke="#9aa0a8" strokeWidth="1.5" strokeDasharray="4 3" markerEnd={`url(#arrow-x-${flow.id})`} />
            <text x={x + 30} y={y + 52} className="flow-edge-label flow-edge-label--muted">{loop.label}</text>
          </g>;
        })()}

        {/* Exception routes drop to a shared "leaves the flow" rail. Each gets
            its own label under its own arrow; a single merged label read as
            though one arrow carried every exception. */}
        {hasException && (() => {
          const exceptions = flow.edges.filter((edge) => edge.kind === "exception");
          const y = FLOW_GEOMETRY.exceptionY;
          const columns = exceptions.map((edge) => flow.lanes.findIndex((lane) => lane[0].id === edge.from));
          const railStart = xFor(Math.min(...columns)) + FLOW_GEOMETRY.boxW / 2;
          const railEnd = xFor(Math.max(...columns)) + FLOW_GEOMETRY.boxW / 2;
          return <g>
            {railEnd > railStart && <line x1={railStart} y1={y} x2={railEnd} y2={y} stroke="#c3c7cd" strokeWidth="1.5" strokeDasharray="4 3" />}
            {exceptions.map((edge, index) => {
              const x = xFor(columns[index]) + FLOW_GEOMETRY.boxW / 2;
              return <g key={edge.from}>
                <path d={`M ${x} ${FLOW_GEOMETRY.top + FLOW_GEOMETRY.boxH} V ${y - 2}`} fill="none" stroke="#9aa0a8" strokeWidth="1.5" strokeDasharray="4 3" markerEnd={`url(#arrow-x-${flow.id})`} />
                <circle cx={x} cy={y} r="2.5" fill="#9aa0a8" />
                <text x={x} y={y + 17} textAnchor="middle" className="flow-edge-label flow-edge-label--muted">{edge.label}</text>
              </g>;
            })}
            <text x={railEnd + 14} y={y + 4} className="flow-edge-label flow-edge-label--muted">leaves the flow</text>
          </g>;
        })()}
      </svg>
    </div>
    <div className="flow-source">Source: {flow.source}</div>
  </figure>;
}

function FlowsPage() {
  return <section className="flows-page">
    <div className="brand-intro"><div><div className="section-kicker">PROCESS FLOWS</div><h2>How work moves through PULSE</h2><p>Four sequences that explain most of the questions people ask: what happens to a travel request, how a document reaches signature, where a ticket goes, and how a hosted tool finds its SharePoint site. Each is drawn from the values the application actually displays.</p></div></div>
    <div className="flow-legend" aria-hidden="true">
      <span className="flow-legend-item flow-legend-item--start">Entry point</span>
      <span className="flow-legend-item flow-legend-item--step">In progress</span>
      <span className="flow-legend-item flow-legend-item--done">Closed</span>
      <span className="flow-legend-item flow-legend-item--exception">Exception</span>
    </div>
    {flows.map((flow) => <FlowDiagram flow={flow} key={flow.id} />)}
    <div className="flow-footnote">A diagram shows the route, not the rules. Who may move a record between these states, and what evidence each move requires, is in the Standard Operating Procedures.</div>
  </section>;
}

function ScreensPage() {
  const total = screenGroups.reduce((sum, group) => sum + group.screens.length, 0);
  return <section className="screens-page">
    <div className="brand-intro"><div><div className="section-kicker">INTERFACE REFERENCE · {SCREEN_CAPTURE_DATE}</div><h2>The current screens</h2><p>{total} captures of the shipping build, grouped the way the app is navigated. Every record shown is synthetic: neutral role display names, example addresses, and no real project or personal data. Use these to confirm a written step matches what a user actually sees.</p></div></div>
    {screenGroups.map((group) => <div className="screen-group" key={group.group}>
      <div className="brand-section-title"><div><div className="section-kicker">{group.group.toUpperCase()}</div><h3>{group.blurb}</h3></div><p>{group.screens.length} {group.screens.length === 1 ? "screen" : "screens"}</p></div>
      <div className="screen-grid">{group.screens.map((screen) => <figure className="screen-card" key={screen.file}>
        <a suppressHydrationWarning href={getAssetUrl(`/screenshots/${screen.file}`)} target="_blank" rel="noreferrer" aria-label={`Open the full-size ${screen.title} capture`}>
          <img suppressHydrationWarning src={getAssetUrl(`/screenshots/${screen.file}`)} alt={`PULSE ${screen.title} screen`} loading="lazy" />
        </a>
        <figcaption><b>{screen.title}</b><code>{screen.route}</code><p>{screen.note}</p></figcaption>
      </figure>)}</div>
    </div>)}
  </section>;
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
    <div className="brand-intro"><div><div className="section-kicker">PULSE BRAND IDENTITY · V1.3</div><h2>The operating rhythm</h2><p>PULSE turns scattered activity into a current, shared, and actionable view of the work. The identity is deliberately restrained: black and white carry the brand, while interface color clarifies action and state.</p></div><div className="brand-downloads"><a suppressHydrationWarning className="download-guide" href={getAssetUrl("/brand-assets/PULSE_Brand_Assets_v1.3.zip")} download>Download all assets <span>↓</span></a><a suppressHydrationWarning href={getAssetUrl("/brand-assets/PULSE_Brand_Identity_Guide_v1.3.pdf")} download>Brand guide PDF</a><a suppressHydrationWarning href={getAssetUrl("/brand-assets/PULSE_Brand_Assets_README.txt")} download>Usage README</a></div></div>
    <div className="brand-principles"><article><span>01</span><h3>Signal in. Action out.</h3><p>The upper dot represents an incoming update or observation. The lower dot completes the pulse: a decision, handoff, approval, or action.</p></article><article><span>02</span><h3>Calm, direct, operational.</h3><p>Use familiar words. State the action, owner, and timing. Be specific without adding drama.</p></article><article><span>03</span><h3>Protect the signal.</h3><p>Keep the mark proportional, clear of nearby elements, and free from effects, status colors, or repositioned dots.</p></article></div>
    <div className="brand-section-title"><div><div className="section-kicker">COLOR</div><h3>Monochrome identity, purposeful interface color</h3></div><p>Use blue for links and current selection; reserve green, amber, and red for status.</p></div>
    <div className="palette"><div className="swatch black"><b>Signal Black</b><span>#070708</span></div><div className="swatch white"><b>Field White</b><span>#FFFFFF</span></div><div className="swatch blue"><b>Command Blue</b><span>#2F66FF</span></div><div className="swatch green"><b>Status Green</b><span>#2B7A5E</span></div><div className="swatch amber"><b>Warning Amber</b><span>#C77800</span></div><div className="swatch red"><b>Critical Red</b><span>#B42318</span></div></div>
    <div className="brand-section-title asset-heading"><div><div className="section-kicker">DOWNLOADS</div><h3>Identity assets</h3></div><p>Use the supplied files as-is. The full guide contains construction, clear-space, and minimum-size rules.</p></div>
    <div className="asset-grid">{assets.map((asset) => <article className={`asset-card ${asset.dark ? "dark" : ""} ${asset.compact ? "compact-asset" : ""}`} key={asset.file}><div className="asset-preview"><img suppressHydrationWarning src={getAssetUrl(`/brand-assets/${asset.file}`)} alt={asset.title} /></div><div><b>{asset.title}</b><p>{asset.note}</p><a suppressHydrationWarning href={getAssetUrl(`/brand-assets/${asset.file}`)} download>Download PNG ↓</a></div></article>)}</div>
  </section>;
}
