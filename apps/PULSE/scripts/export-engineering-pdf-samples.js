#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const ROOT = path.resolve(__dirname, "..");
const TRAVEL_JS = path.join(ROOT, "assets/js/pages/travel.js");
const OUT_DIR = path.join(ROOT, "validation", "pdf-samples");

function loadPdfBuilder() {
  const source = fs.readFileSync(TRAVEL_JS, "utf8");
  const start = source.indexOf("function pdfSafeText");
  const end = source.indexOf("function createSimpleTravelPdfBlob");
  if (start < 0 || end < 0) throw new Error("Could not locate engineering PDF functions in travel.js");
  const block = source.slice(start, end);

  class Blob {
    constructor(parts, opts) {
      this._buf = Buffer.concat(parts.map((part) => (Buffer.isBuffer(part) ? part : Buffer.from(part))));
      this.type = opts && opts.type;
    }
    async arrayBuffer() {
      return this._buf;
    }
  }

  const sandbox = { Blob, console };
  vm.createContext(sandbox);
  vm.runInContext(`${block}; this.createEngineeringTravelPdfBlob = createEngineeringTravelPdfBlob;`, sandbox);
  return sandbox.createEngineeringTravelPdfBlob;
}

const sampleRequest = {
  id: "TR-2026-0142",
  requester: "Scalere, Elijah T CTR (USA)",
  requesterEmail: "elijah.scalere@example.mil",
  destination: "CLIP",
  start: "2026-07-14",
  end: "2026-07-18",
  purpose: "Systems integration support for quarterly release",
  chargeObject: "EU00001",
  formMode: "Engineering",
  engineeringForm: {
    dateOfRequest: "2026-07-01",
    travelerName: "Scalere, Elijah T CTR (USA)",
    phoneNumber: "555-0100",
    purposeOfTdy: "On-site integration testing and stakeholder review",
    onBase: "Yes",
    tdyLocation: "CLIP",
    tdyStartDate: "2026-07-14",
    tdyReturnDate: "2026-07-18",
    numberOfDays: "5",
    flying: "Yes",
    seatPreference: "Aisle",
    flightFrom: "DEN 06:15",
    flightTo: "LAX 09:40",
    flightReturn: "LAX 17:05",
    airportTransport: { mode: "Rental", upgradeJustification: "Required equipment transport", vehicleType: "SUV" },
    tdyTransport: { mode: "Rental", upgradeJustification: "", vehicleType: "Sedan" },
    lodging: [
      { location: "Colorado Springs, CO", checkInDate: "2026-07-14", checkOutDate: "2026-07-16", option1: "Marriott near gate", option2: "On-base lodging" },
      { location: "Los Angeles, CA", checkInDate: "2026-07-16", checkOutDate: "2026-07-18", option1: "Airport hotel", option2: "Downtown option" },
      { location: "", checkInDate: "", checkOutDate: "", option1: "", option2: "" }
    ],
    comments: "Travel aligns with sprint demo week. Rental required for test hardware.",
    bfm: "BFM-42"
  }
};

async function main() {
  const createEngineeringTravelPdfBlob = loadPdfBuilder();
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const variants = [
    { name: "sample-filled", request: sampleRequest },
    {
      name: "sample-minimal",
      request: {
        ...sampleRequest,
        id: "TR-2026-MIN",
        engineeringForm: {
          ...sampleRequest.engineeringForm,
          lodging: [{}, {}, {}],
          comments: "",
          bfm: ""
        }
      }
    }
  ];

  for (const variant of variants) {
    const blob = createEngineeringTravelPdfBlob(variant.request);
    const buf = Buffer.from(await blob.arrayBuffer());
    const outPath = path.join(OUT_DIR, `engineering-${variant.name}.pdf`);
    fs.writeFileSync(outPath, buf);
    console.log(`Wrote ${outPath} (${buf.length} bytes)`);
  }

  console.log("Layout validation passed for all samples.");
}

main().catch((err) => {
  console.error(err.stack || err.message || err);
  process.exit(1);
});
