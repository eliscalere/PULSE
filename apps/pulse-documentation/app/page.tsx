"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { DocumentPage, parsePage } from "./document-reader";
import focusedFigures from "./generated/figures/09_PULSE_Focused_Tools_and_Package_Delivery.json";
import devscaleFigures from "./generated/figures/10_PULSE_Development_Environment_and_Scale_Validation.json";
import interfaceFigures from "./generated/figures/11_PULSE_Interface_Reference.json";
import flowFigures from "./generated/figures/12_PULSE_Process_Flows.json";

/* Figures live in the documents, not in a separate gallery. The text extraction
   the reader renders cannot carry an image, so the generator emits a manifest
   keyed by section and the reader places the same figures the PDF shows. Both
   come from the document source, so neither can drift from the other. */
type Figure = { file: string; caption: string; meta?: string; description?: string; hideCaption?: boolean };
const figuresByDocument: Record<string, Record<string, Figure[]>> = {
  focused: focusedFigures as Record<string, Figure[]>,
  devscale: devscaleFigures as Record<string, Figure[]>,
  interface: interfaceFigures as Record<string, Figure[]>,
  flows: flowFigures as Record<string, Figure[]>,
};

/* Clicking a figure opens it larger in place. A new tab would show the raw
   data: URI with no caption and no way back, which is worse than not zooming. */
function FigureCard({ figure, onOpen }: { figure: Figure; onOpen: (figure: Figure) => void }) {
  const isDiagram = figure.file.endsWith(".svg");
  return <figure className={isDiagram ? "doc-figure doc-figure--diagram" : "doc-figure"}>
    <button type="button" className="doc-figure-open" onClick={() => onOpen(figure)} aria-label={`Enlarge ${figure.caption}`}>
      <img suppressHydrationWarning src={getAssetUrl(figure.file)} alt={figure.caption} loading="lazy" />
      <span className="doc-figure-zoom" aria-hidden="true">Enlarge</span>
    </button>
    <figcaption>
      {!figure.hideCaption && <b>{figure.caption}</b>}
      {figure.meta && <code>{figure.meta}</code>}
      {figure.description && <p>{figure.description}</p>}
    </figcaption>
  </figure>;
}

function FigureLightbox({ figure, onClose }: { figure: Figure; onClose: () => void }) {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => { if (event.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKeyDown);
    /* Keep the page from scrolling behind the overlay. */
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKeyDown);
      document.body.style.overflow = previousOverflow;
    };
  }, [onClose]);

  return <div className="figure-lightbox" role="dialog" aria-modal="true" aria-label={figure.caption}>
    <button type="button" className="figure-lightbox-backdrop" onClick={onClose} aria-label="Close" tabIndex={-1} />
    <div className="figure-lightbox-panel">
      <header>
        <div>
          <b>{figure.caption}</b>
          {figure.meta && <code>{figure.meta}</code>}
        </div>
        <button type="button" className="figure-lightbox-close" onClick={onClose} autoFocus aria-label="Close the enlarged figure">Close ✕</button>
      </header>
      <div className="figure-lightbox-image">
        <img suppressHydrationWarning src={getAssetUrl(figure.file)} alt={figure.caption} />
      </div>
      {figure.description && <p>{figure.description}</p>}
    </div>
  </div>;
}

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
  { id: "interface", number: "11", title: "Interface Reference", type: "Visual reference", audience: "All PULSE users", pages: 7, description: "Eighteen captures of the shipping build, grouped the way the application is navigated, for checking that a written step matches what a user actually sees.", topics: ["Interface", "Screens", "Reference"], textFile: "11_PULSE_Interface_Reference.txt", pdfFile: "11_PULSE_Interface_Reference.pdf" },
  { id: "flows", number: "12", title: "Process Flows", type: "Visual reference", audience: "All PULSE users", pages: 6, description: "How a travel request, a document review, a support ticket, and site resolution in a web part each move from entry point to closed state, drawn from the values the application displays.", topics: ["Flows", "Travel", "Review"], textFile: "12_PULSE_Process_Flows.txt", pdfFile: "12_PULSE_Process_Flows.pdf" },
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

/* The default package omits the controlled PDFs to stay fast (see the packager).
   Anything that links to one has to check first, or it renders a dead link that
   resolves against the SharePoint page and returns the site's own HTML. On the
   dev server, where assets are served normally, they are always available. */
/* A data: URI opened in a tab has no filename and no way back. Converting it to
   a Blob lets the browser save it as the controlled document's own name. */
function downloadPdf(doc: Document) {
  const uri = getAssetUrl(`/source-pdfs/${doc.pdfFile}`);
  if (!uri.startsWith("data:")) { window.open(uri, "_blank", "noreferrer"); return; }
  const base64 = uri.slice(uri.indexOf(",") + 1);
  const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = doc.pdfFile;
  document.body.appendChild(link);
  link.click();
  link.remove();
  /* Revoke on the next tick; revoking synchronously cancels the download. */
  window.setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function hasAsset(url: string): boolean {
  if (typeof window === "undefined") return false;
  const assets = (window as unknown as { __PULSE_ASSETS__?: Record<string, string> }).__PULSE_ASSETS__;
  if (!assets) return true; // dev server: real files behind real paths
  return typeof assets[url] === "string";
}

export default function Home() {
  const [active, setActive] = useState("overview");
  const [view, setView] = useState<"docs" | "brand">("docs");
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
  /* Starts true so the first client render matches the server markup, then the
     boot effect corrects it. Rendering it straight from hasAsset() would be a
     hydration mismatch, since the server cannot see the asset map. */
  const [pdfsBundled, setPdfsBundled] = useState(true);
  const [lightbox, setLightbox] = useState<Figure | null>(null);

  useEffect(() => {
    setPdfsBundled(hasAsset(`/source-pdfs/${documents[1].pdfFile}`));
  }, []);

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
  function search(value: string) { setQuery(value); setIsSearching(Boolean(value.trim())); }
  function goToSection(page: number) {
    setMobileOpen(false);
    setActiveSection(page);
    scrollToSection(page);
  }

  return <>
    {bootPhase !== "ready" && <PulseBootLoader exiting={bootPhase === "exiting"} />}
    {lightbox && <FigureLightbox figure={lightbox} onClose={() => setLightbox(null)} />}
    <main className={bootPhase === "ready" ? "pulse-app-shell" : "pulse-app-shell pulse-app-shell--booting"} aria-hidden={bootPhase === "ready" ? undefined : true}>
    <aside className={`sidebar ${mobileOpen ? "open" : ""}`}>
      <a className="brand" href="#top" onClick={() => choose("overview")}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Wordmark_White_Transparent.png")} alt="PULSE" /></a>
      <div className="side-label">Documentation</div>
      <nav className="document-nav">{documents.filter((doc) => doc.id !== "brand").map((doc) => {
        const isActive = active === doc.id && !isSearching && view === "docs";
        return <div className="nav-group" key={doc.id}><button className={isActive ? "nav-item active" : "nav-item"} onClick={() => choose(doc.id)}><span>{doc.number}</span><b>{doc.title}</b></button>{isActive && selectedSections.length > 0 && <div className="subnav" aria-label={`${doc.title} sections`}>{selectedSections.map((section) => <button className={activeSection === section.page ? "subnav-item active" : "subnav-item"} aria-current={activeSection === section.page ? "true" : undefined} key={section.page} onClick={() => goToSection(section.page)}><span>{String(section.page).padStart(2, "0")}</span><b>{section.title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "")}</b></button>)}</div>}</div>;
      })}</nav>
      <div className="side-footer"><span className="status-dot" /> Complete source library<br /><small>{documents.reduce((sum, doc) => sum + doc.pages, 0)} pages · July 2026</small></div>
      <button className={view === "brand" ? "brand-button active" : "brand-button"} onClick={openBrand}><img suppressHydrationWarning src={getAssetUrl("/brand-assets/PULSE_Dot_Mark_White_Transparent.png")} alt="" /><span><b>Brand guidelines</b><small>Identity & assets</small></span><strong>→</strong></button>
    </aside>
    <section className="content" id="top">
      <header className="topbar"><button className="menu" aria-label="Open documentation navigation" onClick={() => setMobileOpen(!mobileOpen)}>☰</button><div className="crumb">PULSE / <span>{isSearching ? "Search" : view === "brand" ? "Brand guidelines" : `${selected.title}${currentSection ? ` / ${currentSection.title.replace(/^\d+(?:\.\d+)?\s*\/\s*/, "")}` : ""}`}</span></div>{view === "brand" ? <a suppressHydrationWarning className="source-link" href={getAssetUrl("/brand-assets/PULSE_Brand_Identity_Guide_v1.3.pdf")} download>Download guide PDF ↓</a> : <button type="button" className="source-link source-link--action" onClick={() => (pdfsBundled ? downloadPdf(selected) : window.print())}>{pdfsBundled ? "Download PDF ↓" : "Print ↓"}</button>}</header>
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
            
          </article>)}
        </div>)}</div>
      </section> : view === "brand" ? <BrandPage /> : <>
        <section className="document-header"><div><div className="section-kicker">DOCUMENT {selected.number} · {selected.type}</div><h2>{selected.title}</h2><p>{selected.description}</p><div className="meta"><span>{selected.pages} pages</span><span>{selected.audience}</span></div></div><div className="doc-actions">
            {pdfsBundled
              ? <button type="button" className="doc-action doc-action--primary" onClick={() => downloadPdf(selected)}><span>PDF</span><b>Download this document</b><small>Controlled PDF · UNCLASSIFIED on every page</small></button>
              : <button type="button" className="doc-action doc-action--primary" onClick={() => window.print()}><span>PRINT</span><b>Print this document</b><small>Full document with figures · marked UNCLASSIFIED</small></button>}
            <button type="button" className="doc-action" onClick={() => window.print()}><span>PRINT</span><b>Print or save a copy</b><small>Uses the browser print dialog</small></button>
          </div></section>
        <section className="reader" id="reader-top">
          <div className="reader-heading"><div><span className="reading-position">{selectedPages.length || selected.pages} sections · continuous</span><h3>Read straight through, or jump from the sidebar</h3></div><div className="topics">{selected.topics.map((topic) => <span key={topic}>{topic}</span>)}</div></div>
          {!selectedPages.length
            ? <div className="reader-skeleton" aria-label="Preparing the structured reading view"><span /><span /><span /><span /></div>
            : <div className="reader-stream">{selectedPages.map((raw, index) => {
              const sectionFigures = figuresByDocument[selected.id]?.[String(index + 1)];
              return <div className="reader-section" data-section={index + 1} id={`section-${index + 1}`} key={`${selected.id}-${index}`}>
                <DocumentPage raw={raw} pageNumber={index + 1} documentTitle={selected.title} figures={sectionFigures?.map((figure) => <FigureCard figure={figure} onOpen={setLightbox} key={figure.file} />)} />
              </div>;
            })}</div>}
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
