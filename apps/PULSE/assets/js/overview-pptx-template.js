/* Normalized from the supplied My Overview and Team Overview OOXML packages.
   This is the runtime presentation contract used by overview-pptx-export.js. */
(function (root) {
  root.PULSE_OVERVIEW_PPTX_TEMPLATE = Object.freeze({
    schemaVersion: "1.0",
    slide: { width: 10, height: 7.5, layout: "LAYOUT_4x3" },
    fontFace: "Inter",
    assets: {
      wordmarkWhite: "assets/images/pulse-wordmark-white.png",
      wordmarkBlack: "assets/images/pulse-wordmark-black.png",
      seal: "assets/images/overview-seal.png"
    },
    colors: {
      ink: "070708",
      white: "FFFFFF",
      accent: "2F66FF",
      muted: "51545A",
      quiet: "73767C",
      line: "DADDE2",
      tableHead: "ECEDEF",
      tableAlt: "F8F9FA",
      barTrack: "ECEDEF",
      coverTag: "121317",
      coverMuted: "C9CCD2",
      coverMeta: "D4D6DA",
      coverQuiet: "AEB2BA",
      coverPanel: "0C0E12",
      coverPanelLine: "373A42",
      green: "2B7A5E",
      greenTint: "EAF5F1",
      amber: "C77800",
      amberTint: "FFF4E2",
      red: "B42318",
      redTint: "FDECEA",
      blueTint: "EAF0FF",
      neutralTint: "F2F3F5"
    },
    chrome: {
      wordmark: { x: 0.36, y: 0.15, w: 1.0, h: 0.33 },
      section: { x: 1.70, y: 0.22, w: 3.0, h: 0.17 },
      seal: { x: 9.05, y: 0.11, w: 0.58, h: 0.58 },
      topRule: { x: 0.36, y: 0.76, w: 9.28, h: 0 },
      title: { x: 0.36, y: 1.0, w: 8.45, h: 0.46 },
      footerRule: { x: 0.36, y: 7.02, w: 9.28, h: 0 },
      footerLeft: { x: 0.36, y: 7.15, w: 4.8, h: 0.16 },
      footerRight: { x: 7.55, y: 7.15, w: 2.10, h: 0.16 }
    },
    cui: {
      topLabel: { x: 0, y: 0.035, w: 10, h: 0.18 },
      topRule: { x: 0.20, y: 0.235, w: 9.60, h: 0 },
      bottomRule: { x: 0.20, y: 7.215, w: 9.60, h: 0 },
      bottomLabel: { x: 0, y: 7.285, w: 10, h: 0.18 }
    },
    cover: {
      wordmark: { x: 0.64, y: 0.58, w: 2.22, h: 0.73 },
      accentRule: { x: 0.66, y: 1.50, w: 0.72, h: 0 },
      title: { x: 0.64, y: 1.86, w: 5.70, h: 0.78 },
      subtitle: { x: 0.66, y: 2.78, w: 5.55, h: 0.40 },
      tag: { x: 0.66, y: 3.65, w: 2.45, h: 0.30 },
      meta: { x: 0.66, y: 5.45, w: 3.90, h: 0.58 },
      seal: { x: 7.12, y: 1.04, w: 2.04, h: 2.04 },
      classificationPanel: { x: 5.55, y: 5.15, w: 3.95, h: 0.92 },
      classificationStripe: { x: 5.55, y: 5.15, w: 0.05, h: 0.92 },
      classificationTitle: { x: 5.71, y: 5.23, w: 3.70, h: 0.16 },
      classificationDetail: { x: 5.71, y: 5.42, w: 3.67, h: 0.51 },
      statementDetail: { x: 5.45, y: 6.23, w: 3.80, h: 0.22 }
    },
    pagination: {
      endItemsPerSlide: 5,
      teamProjectsPerSlide: 9,
      myProjectsPerSlide: 5,
      tasksPerSlide: 9,
      workloadPerSlide: 7,
      projectTasks: 3,
      projectRisks: 3
    }
  });
}(window));
