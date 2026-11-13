function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 DCP Sync")
    .addItem("🔄 Update Status (Completed & Pending Docs)", "syncDCPDeliveryStatus")
    .addSeparator()
    .addItem("⏰ Enable Auto-Sync (Every 15 Mins)", "createAutoSyncTrigger")
    .addToUi();
}

function syncDCPDeliveryStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw_Delivery");
  const targetTabs = ["2025c", "2026"];

  // 1. Build lookup from Raw_Delivery (TRACES)
  const bySidLot = {};
  const bySid = {};

  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawVals = rawSheet.getDataRange().getDisplayValues();
    const headers = rawVals[0].map(h => String(h || "").toLowerCase().trim());

    let siteCol = findColIndex(headers, ["site code", "school id", "school_id", "beis"]);
    if (siteCol === -1 && headers.length >= 6) siteCol = 5; // Col F fallback

    let statCol = findColIndex(headers, ["current status", "delivery status", "status"]);
    if (statCol === -1 && headers.length >= 7) statCol = 6; // Col G fallback

    const lotCol = findColIndex(headers, ["lot"]);

    if (siteCol !== -1 && statCol !== -1) {
      for (let r = 1; r < rawVals.length; r++) {
        const sid = String(rawVals[r][siteCol] || "").trim();
        const st = String(rawVals[r][statCol] || "").trim();
        const lot = lotCol !== -1 ? String(rawVals[r][lotCol] || "").trim() : "";

        if (!sid || !st) continue;
        const prio = getStatusPriority(st);

        if (!bySid[sid] || prio >= bySid[sid].prio) {
          bySid[sid] = { status: st, prio: prio };
        }
        if (lot) {
          const key = sid + "|" + lot;
          if (!bySidLot[key] || prio >= bySidLot[key].prio) {
            bySidLot[key] = { status: st, prio: prio };
          }
        }
      }
    }
  }

  let updatedCompleted = 0;
  let updatedPendingDocs = 0;

  // 2. Update Column O (Status) in 2025c and 2026
  targetTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const lastRow = sheet.getLastRow();
    const numRows = lastRow - 1;
    const maxCols = Math.max(sheet.getLastColumn(), 15);

    const displayVals = sheet.getRange(2, 1, numRows, maxCols).getDisplayValues();
    const formulasColO = sheet.getRange(2, 15, numRows, 1).getFormulas();
    const outColO = [];

    for (let i = 0; i < numRows; i++) {
      const row = displayVals[i];
      const pkg = String(row[2] || "").trim();      // Col C: DCP Package
      const currentColO = String(row[14] || "").trim(); // Col O: Status

      // Skip blank rows or rows where Col O has an existing formula
      if (!pkg || formulasColO[i][0]) {
        outColO.push([formulasColO[i][0] ? formulasColO[i][0] : currentColO]);
        continue;
      }

      const lot = String(row[0] || "").trim();       // Col A: LOT
      const sid = String(row[4] || "").trim();       // Col E: School ID
      const drDate = row[10];                        // Col K: DR Date
      const iarDate = row[11];                       // Col L: IAR Date
      const colN = String(row[13] || "").trim();     // Col N: TRACES Status

      const hasDR = isValidEntry(drDate);
      const hasIAR = isValidEntry(iarDate);

      // Highest priority status between Raw_Delivery (SID+LOT), Raw_Delivery (SID), and Col N
      const candLot = (sid && lot && bySidLot[sid + "|" + lot]) ? bySidLot[sid + "|" + lot] : { status: "", prio: 0 };
      const candSid = (sid && bySid[sid]) ? bySid[sid] : { status: "", prio: 0 };
      const candColN = { status: colN, prio: getStatusPriority(colN) };

      let bestRaw = candColN;
      if (candSid.prio >= bestRaw.prio) bestRaw = candSid;
      if (candLot.prio >= bestRaw.prio) bestRaw = candLot;

      const rawLow = bestRaw.status.toLowerCase();
      const origLow = currentColO.toLowerCase();

      const isCompleted =
        (hasDR && hasIAR) ||
        rawLow.includes("complete") ||
        origLow.includes("complete");

      const isDeliveredPendingDocs =
        !isCompleted &&
        (hasDR ||
          hasIAR ||
          /delivered|pending mov|pending dr|pending iar/.test(rawLow) ||
          /delivered|pending mov|pending dr|pending iar/.test(origLow));

      let newStatus = currentColO;
      if (isCompleted) {
        newStatus = "Completed";
        if (currentColO !== "Completed") updatedCompleted++;
      } else if (isDeliveredPendingDocs) {
        newStatus = "Delivered (Pending Docs)";
        if (currentColO !== "Delivered (Pending Docs)") updatedPendingDocs++;
      }

      outColO.push([newStatus]);
    }

    sheet.getRange(2, 15, numRows, 1).setValues(outColO);
  });

  try {
    SpreadsheetApp.getUi().alert(
      "✅ DCP Status Sync Complete!\n\n" +
      "• Updated to Completed: " + updatedCompleted + " rows\n" +
      "• Updated to Delivered (Pending Docs): " + updatedPendingDocs + " rows\n\n" +
      "Pindutin ang 'Sync Latest Google Sheet' sa Streamlit dashboard para mag-reflect agad."
    );
  } catch (e) {
    // Runs silently when triggered automatically in the background
  }
}

function createAutoSyncTrigger() {
  // Remove duplicate triggers first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "syncDCPDeliveryStatus") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("syncDCPDeliveryStatus")
    .timeBased()
    .everyMinutes(15)
    .create();

  SpreadsheetApp.getUi().alert(
    "⏰ Auto-Sync Enabled!\n\nAwtomatiko nang mag-a-update ang 'Completed' at 'Delivered (Pending Docs)' sa Google Sheet tuwing 15 minuto."
  );
}

function findColIndex(headers, keywords) {
  for (let c = 0; c < headers.length; c++) {
    if (keywords.some(k => headers[c].includes(k))) return c;
  }
  return -1;
}

function isValidEntry(val) {
  if (val instanceof Date) return true;
  const s = String(val || "").trim().toLowerCase();
  const invalid = new Set([
    "", "-", "--", "n/a", "na", "none", "null", "tba",
    "pending", "incomplete", "for iar", "for dr", "for submission",
    "no", "0", "nan", "nat"
  ]);
  return !invalid.has(s);
}

function getStatusPriority(statusText) {
  const low = String(statusText || "").toLowerCase().trim();
  if (!low) return 0;
  if (low.includes("complete")) return 3;
  if (/delivered|pending mov|pending dr|pending iar/.test(low)) return 2;
  return 1;
}
