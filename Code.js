// === CONFIG ===
const TEMPLATE_ID = "1bss96EkFlNl_GQ35E8PdxeZrgO-02yrJILtevKzE4SE"; // Google Doc template ID
const FOLDER_ID = "1p-nYwiyl6XsXXV93_nJf5OoeXXmaWZWL"; // Folder for PDFs
const SHEET_FILE_ID = "1MVY1ucbqCTRQkoEEMaQc6tEI6u62psbup6iL023xGsI"; // Spreadsheet ID
const REDIRECT_URL = "https://sites.google.com/view/giantmotoprocorp";

// === FRONTEND LOADER ===
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Requisition Form v15")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === MAIN SUBMISSION HANDLER (Optimized for Speed) ===
function saveAndCreatePdf(data) {
  if (!data || typeof data !== "object") throw new Error("Invalid data.");
  if (!Array.isArray(data.items)) data.items = [];

  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sheetName = getSheetNameByPurpose(data.purpose, data.branch);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Create header only if new sheet
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp",
      "Branch",
      "Date",
      "To",
      "Purpose",
      "ITEM_ID",
      "Qty",
      "Unit",
      "Description",
      "UPrice",
      "Amount",
      "Total",
      "Requested By",
      "Status",
      "Release Date",
      "Received By",
      "PDF URL",
    ]);
  }

  const ts = new Date();
  const pdfUrl = createPdfFromTemplate(data, ts);

  // Prepare fast bulk append
  const newRows = data.items.map((it) => {
    const itemId = findItemId(it.description);
    return [
      ts,
      data.branch,
      data.date,
      data.to,
      data.purpose,
      itemId,
      it.qty,
      it.unit,
      it.description,
      it.uprice,
      it.amount,
      data.total,
      data.requested_by,
      "Pending",
      "",
      "",
      pdfUrl,
    ];
  });

  // Write all items in one call (faster)
  if (newRows.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  // Try to schedule SUMMARY generation (non-blocking)
  try {
    ScriptApp.newTrigger("generateMasterSummary")
      .timeBased()
      .after(1000)
      .create();
  } catch (e) {
    // ignore trigger creation errors
  }

  return pdfUrl;
}

// === GET SHEET NAME BY PURPOSE/BRANCH ===
function getSheetNameByPurpose(purpose, branch) {
  const p = purpose ? purpose.toUpperCase() : "";
  if (p === "SPECIAL REQUEST") return "SPECIAL REQUEST";
  return branch ? branch.trim().toUpperCase() : "GENERAL";
}

// === FIND ITEM ID (FROM REAL-TIME STOCKS) ===
function findItemId(description) {
  if (!description) return "";
  const d = description.toString().trim().toUpperCase();
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("REAL-TIME STOCKS");
  if (!sh) return "";

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "";

  // Search column B (Description) for a match, return column A (Item_ID)
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    const item_id = (data[i][0] || "").toString().trim();
    const desc = (data[i][1] || "").toString().trim().toUpperCase();
    if (desc === d) return item_id;
  }
  return "";
}

// === FAST PDF CREATION (no external API, lightweight) ===
function createPdfFromTemplate(data, timestamp) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const copy = templateFile.makeCopy(
    `Requisition ${Utilities.formatDate(
      timestamp,
      "Asia/Manila",
      "yyyy-MM-dd HH:mm:ss"
    )}`
  );

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  try {
    body.replaceText("{{BRANCH}}", data.branch || "");
    body.replaceText("{{DATE}}", data.date || "");
    body.replaceText("{{TO}}", data.to || "");
    body.replaceText("{{PURPOSE}}", data.purpose || "");
    body.replaceText("{{TOTAL}}", data.total || "");
    body.replaceText("{{NOTES}}", data.note || "");
    body.replaceText("{{REQUESTED_BY}}", data.requested_by || "");

    // Insert items table
    const search = body.findText("{{ITEMS}}");
    if (search) {
      const element = search.getElement();
      const parent = element.getParent();

      const tableData = [["Qty", "Unit", "Description", "UPrice", "Amount"]];
      (data.items || []).forEach((it) => {
        tableData.push([
          it.qty || "",
          it.unit || "",
          it.description || "",
          it.uprice || "",
          it.amount || "",
        ]);
      });
      tableData.push(["", "", "", "Total", data.total || ""]);

      const table = body.insertTable(body.getChildIndex(parent) + 1, tableData);
      table.setBorderWidth(1);
      table.setFontSize(9);

      for (let i = 0; i < table.getNumRows(); i++) {
        const row = table.getRow(i);
        for (let j = 0; j < row.getNumCells(); j++) {
          const cell = row.getCell(j);
          cell
            .setPaddingTop(3)
            .setPaddingBottom(3)
            .setPaddingLeft(5)
            .setPaddingRight(5);
          try {
            if (j === 0) cell.setWidth(40);
            if (j === 1) cell.setWidth(60);
            if (j === 2) cell.setWidth(240);
            if (j === 3) cell.setWidth(80);
            if (j === 4) cell.setWidth(80);
          } catch (e) {}
          const text = cell.getChild(0);
          if (
            text &&
            text.getType &&
            text.getType() === DocumentApp.ElementType.PARAGRAPH
          ) {
            const para = text.asParagraph();
            if (i === 0) {
              para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              para.setBold(true);
              try {
                cell.setBackgroundColor("#f0f0f0");
              } catch (e) {}
            } else if ([0, 1, 3, 4].includes(j)) {
              para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            } else {
              para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
            }
          }
        }
      }
      element.asText().setText("");
    }

    // Signature insertion
    if (data.requested_by_signature) {
      const parts = data.requested_by_signature.split(",");
      const base64Signature = parts.length > 1 ? parts[1] : parts[0];
      const base64Data = base64Signature.trim();
      if (base64Data) {
        const imgBytes = Utilities.base64Decode(
          base64Data.replace(/^data:image\/png;base64,/, "")
        );
        const sigBlob = Utilities.newBlob(
          imgBytes,
          "image/png",
          "signature.png"
        );
        const searchSig = body.findText("{{REQUESTED_BY_SIGNATURE}}");
        if (searchSig) {
          const el = searchSig.getElement();
          const parent = el.getParent();
          const insertIndex = parent.getChildIndex(el);
          el.asText().setText("");
          const spacerBlob = Utilities.newBlob(
            Utilities.base64Decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQMAAAAl21bKAAAAA1BMVEUAAACnej3aAAAAAXRSTlMAQObYZgAAAApJREFUCNdjYAAAAAIAAeIhvDMAAAAASUVORK5CYII="
            ),
            "image/png",
            "spacer.png"
          );
          parent
            .insertInlineImage(insertIndex, spacerBlob)
            .setWidth(95)
            .setHeight(8);
          parent
            .insertInlineImage(insertIndex + 1, sigBlob)
            .setWidth(150)
            .setHeight(55);
        }
      }
    }

    doc.saveAndClose();
  } catch (e) {
    Logger.log("PDF generation error: " + e);
  }

  const pdfBlob = copy.getBlob().getAs("application/pdf");
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob).setName(copy.getName() + ".pdf");
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // Trash temp doc
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdfFile.getUrl();
}
// === REAL-TIME STOCK FETCHER (based on REAL-TIME STOCKS SHEET) ===
function getRunningStocks() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("REAL-TIME STOCKS");
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  // Columns:
  // A = Item_ID
  // B = Description
  // C = Real-Time Stocks
  const data = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  const stockMap = {};

  data.forEach((r) => {
    const description = (r[1] || "").toString().trim();
    const stock = Number(r[2]) || 0;
    if (description) {
      stockMap[description] = stock;
    }
  });

  return stockMap;
}

// === FULL REAL-TIME STOCKS (with unit) ===
function getRealTimeStocksFull() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("REAL-TIME STOCKS");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Read columns A=Item_ID, B=Description, C=Real-Time Stocks, D=Unit
  const data = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  const rows = [];
  data.forEach((r) => {
    const item_id = (r[0] || "").toString().trim();
    const description = (r[1] || "").toString().trim();
    const stock = Number(r[2]) || 0;
    const unit = (r[3] || "").toString().trim().toUpperCase();
    if (description) {
      rows.push({ item_id, description, stock, unit });
    }
  });
  return rows;
}

// === BRANCH LIST FETCHER ===
function getBranchList() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName("Branch List");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Branch header is in column B (index 2). Read from row 2, column 2.
  const values = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  const branches = [];
  const seen = {};
  values.forEach((r) => {
    const v = (r[0] || "").toString().trim();
    if (v && !seen[v]) {
      branches.push(v);
      seen[v] = true;
    }
  });
  return branches;
}
