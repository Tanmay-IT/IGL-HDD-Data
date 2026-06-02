/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import { JWT } from "google-auth-library";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json({ limit: "10mb" }));

// Initialize Gemini SDK with telemetry User-Agent
const ai = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

// Define type models
interface PartitionItem {
  id: string;
  itemType: "Folder" | "File" | "Outlook PST" | "Backup" | "Other";
  name: string;
  formatExtension: string;
  freeSpaceAvailable: string;
  notes: string;
}

interface PartitionBlock {
  id: string;
  driveLetter: string;
  driveType: "System" | "Data" | "Backup" | "Other";
  items: PartitionItem[];
}

interface HDDInventoryItem {
  srNo: number;
  allottedTo: string;
  brand: string;
  type: string;
  capacity: string;
  serialNo: string;
  dateAllotted: string;
  informationUrl: string;
  partitions: PartitionBlock[];
}

// In-Memory Persistent Mock Database for instant sandbox play (Fallback)
let inMemoryDatabase: HDDInventoryItem[] = [
  {
    srNo: 1,
    allottedTo: "Swapnil Shinde",
    brand: "Western Digital",
    type: "SSD",
    capacity: "512 GB",
    serialNo: "HDD-W2A89S",
    dateAllotted: "2024-05-15",
    informationUrl: "#subsheet-HDD-W2A89S",
    partitions: [
      {
        id: "p1",
        driveLetter: "C:\\",
        driveType: "System",
        items: [
          {
            id: "i1",
            itemType: "Folder",
            name: "Windows/System32",
            formatExtension: "System Files",
            freeSpaceAvailable: "85.2 of 146 GB",
            notes: "OS Installation and essential drivers",
          },
          {
            id: "i2",
            itemType: "Outlook PST",
            name: "corporate_archive_2024",
            formatExtension: "pst",
            freeSpaceAvailable: "12.4 of 50 GB",
            notes: "Swapnil Shinde's primary corporate mailbox archive from mid-2024",
          },
        ],
      },
      {
        id: "p2",
        driveLetter: "D:\\",
        driveType: "Data",
        items: [
          {
            id: "i3",
            itemType: "Backup",
            name: "source_code_retail_branch",
            formatExtension: "zip",
            freeSpaceAvailable: "220 of 300 GB",
            notes: "Retail core banking branch application back-up",
          },
        ],
      },
    ],
  },
  {
    srNo: 2,
    allottedTo: "Anjali Gupta",
    brand: "Seagate",
    type: "HDD",
    capacity: "2 TB",
    serialNo: "HDD-S8D90X",
    dateAllotted: "2024-11-20",
    informationUrl: "#subsheet-HDD-S8D90X",
    partitions: [
      {
        id: "p3",
        driveLetter: "E:\\",
        driveType: "Data",
        items: [
          {
            id: "i4",
            itemType: "File",
            name: "swapnil_migration_manifest",
            formatExtension: "csv",
            freeSpaceAvailable: "850.4 of 1000 GB",
            notes: "All corporate devices migration matrix spreadsheet",
          },
          {
            id: "i5",
            itemType: "Backup",
            name: "Outlook_Anjali_Backup_2024",
            formatExtension: "pst",
            freeSpaceAvailable: "420 of 1000 GB",
            notes: "Outlook email folder copy including attachments up to Q3 2024",
          },
        ],
      },
    ],
  },
  {
    srNo: 3,
    allottedTo: "John Doe",
    brand: "Crucial",
    type: "NVMe",
    capacity: "1 TB",
    serialNo: "HDD-C4F56A",
    dateAllotted: "2025-02-14",
    informationUrl: "#subsheet-HDD-C4F56A",
    partitions: [
      {
        id: "p4",
        driveLetter: "C:\\",
        driveType: "System",
        items: [
          {
            id: "i6",
            itemType: "Folder",
            name: "Users/John/Documents",
            formatExtension: "Directories",
            freeSpaceAvailable: "50 of 400 GB",
            notes: "Local work documents & profiles",
          },
        ],
      },
    ],
  },
];

// Stores current dynamic Sheets credentials in server memory for active session
let sheetCredentials = {
  spreadsheetId: "",
  clientEmail: "",
  privateKey: "",
  appsScriptUrl: "", // Optional Google Apps Script URL for direct zero-GCP setup
  useSampleDatabase: true,
};

// Interface for Audit Logging
interface AuditLogItem {
  timestamp: string;
  user: string;
  action: string;
  details: string;
}

let inMemoryAuditLogs: AuditLogItem[] = [
  {
    timestamp: "2026-06-02 07:27:15",
    user: "System Admin",
    action: "Init System",
    details: "Smart Corporate HDD Inventory and Partition Manager successfully loaded."
  }
];

// Ensure AuditLog tab exists
async function ensureAuditLogSheetExists(spreadsheetId: string, accessToken: string) {
  try {
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;
    const res = await fetch(metaUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!res.ok) return;
    const meta = await res.json();
    const exists = meta.sheets?.some((s: any) => s.properties.title === "AuditLog");
    if (!exists) {
      await updateSpreadsheet(spreadsheetId, {
        requests: [
          {
            addSheet: {
              properties: {
                title: "AuditLog",
                gridProperties: { rowCount: 1000, columnCount: 4 }
              }
            }
          }
        ]
      }, accessToken);
      // Write Header
      await writeSheetValues(spreadsheetId, "AuditLog!A1:D1", [["Timestamp", "User", "Action Performed", "Details of Change"]], accessToken);
    }
  } catch (err) {
    console.error("Failed to ensure AuditLog sheets tab:", err);
  }
}

// Log an audit trail entry
async function addAuditLog(action: string, details: string, user: string = "System Admin") {
  const timestamp = new Date().toISOString().replace("T", " ").substring(0, 19);
  const cleanUser = user || "System Admin";
  const item: AuditLogItem = { timestamp, user: cleanUser, action, details };
  inMemoryAuditLogs.unshift(item);

  // If live Google Apps Script web app is connected
  if (sheetCredentials.appsScriptUrl) {
    try {
      await fetch(sheetCredentials.appsScriptUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "audit", timestamp, user: cleanUser, actionName: action, details }),
      });
    } catch (err) {
      console.warn("Apps Script audit log fail:", err);
    }
  } else if (!sheetCredentials.useSampleDatabase && sheetCredentials.spreadsheetId) {
    try {
      const token = await getAuthToken();
      if (token) {
        const ssId = sheetCredentials.spreadsheetId;
        await ensureAuditLogSheetExists(ssId, token);
        const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${ssId}/values/AuditLog!A1:D1:append?valueInputOption=USER_ENTERED`;
        await fetch(appendUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ values: [[timestamp, cleanUser, action, details]] }),
        });
      }
    } catch (err) {
      console.warn("Google Sheet REST audit log fail:", err);
    }
  }
}

// Internal helper: Secure OAuth token retriever for Google APIs using the provided keys
async function getAuthToken() {
  if (
    sheetCredentials.useSampleDatabase ||
    !sheetCredentials.clientEmail ||
    !sheetCredentials.privateKey
  ) {
    return null;
  }
  try {
    const formattedKey = sheetCredentials.privateKey.replace(/\\n/g, "\n");
    const jwt = new JWT({
      email: sheetCredentials.clientEmail,
      key: formattedKey,
      scopes: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
      ],
    });
    const result = await jwt.authorize();
    return result.access_token;
  } catch (error) {
    console.error("JWT Authentication failed:", error);
    throw new Error("Google API Credentials Verification Failed. Please check client_email and private_key.");
  }
}

// REST helper to fetch sheet values
async function fetchSheetValues(spreadsheetId: string, range: string, accessToken: string) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Google Sheets API returns ${res.status}: ${text}`);
  }
  return res.json();
}

// REST helper to batch update spreadsheet
async function updateSpreadsheet(spreadsheetId: string, body: any, accessToken: string) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`BatchUpdate returns ${res.status}: ${text}`);
  }
  return res.json();
}

// REST helper to update range values
async function writeSheetValues(
  spreadsheetId: string,
  range: string,
  values: any[][],
  accessToken: string
) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values }),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`WriteValues returns ${res.status}: ${text}`);
  }
  return res.json();
}

// --- API ROUTES ---

// Configure sheets connection secrets
app.post("/api/auth/save-credentials", (req, res) => {
  const { spreadsheetId, clientEmail, privateKey, appsScriptUrl, useSampleDatabase } = req.body;
  sheetCredentials = {
    spreadsheetId: spreadsheetId || "",
    clientEmail: clientEmail || "",
    privateKey: privateKey || "",
    appsScriptUrl: appsScriptUrl || "",
    useSampleDatabase: useSampleDatabase !== false,
  };

  res.json({
    success: true,
    message: sheetCredentials.useSampleDatabase
      ? "Using Local Persistent Simulation Sandbox Database."
      : sheetCredentials.appsScriptUrl
      ? "Google Apps Script Web App configured successfully!"
      : "Google Sheets Connection Info updated in server-side session memory.",
  });
});

// Resets/reloads local simulation template state
app.post("/api/auth/reset-mock", (req, res) => {
  inMemoryDatabase = [
    {
      srNo: 1,
      allottedTo: "Swapnil Shinde",
      brand: "Western Digital",
      type: "SSD",
      capacity: "512 GB",
      serialNo: "HDD-W2A89S",
      dateAllotted: "2024-05-15",
      informationUrl: "#subsheet-HDD-W2A89S",
      partitions: [
        {
          id: "p1",
          driveLetter: "C:\\",
          driveType: "System",
          items: [
            {
              id: "i1",
              itemType: "Folder",
              name: "Windows/System32",
              formatExtension: "System Files",
              freeSpaceAvailable: "85.2 of 146 GB",
              notes: "OS Installation and essential drivers",
            },
            {
              id: "i2",
              itemType: "Outlook PST",
              name: "corporate_archive_2024",
              formatExtension: "pst",
              freeSpaceAvailable: "12.4 of 50 GB",
              notes: "Swapnil Shinde's primary corporate mailbox archive from mid-2024",
            },
          ],
        },
        {
          id: "p2",
          driveLetter: "D:\\",
          driveType: "Data",
          items: [
            {
              id: "i3",
              itemType: "Backup",
              name: "source_code_retail_branch",
              formatExtension: "zip",
              freeSpaceAvailable: "220 of 300 GB",
              notes: "Retail core banking branch application back-up",
            },
          ],
        },
      ],
    },
    {
      srNo: 2,
      allottedTo: "Anjali Gupta",
      brand: "Seagate",
      type: "HDD",
      capacity: "2 TB",
      serialNo: "HDD-S8D90X",
      dateAllotted: "2024-11-20",
      informationUrl: "#subsheet-HDD-S8D90X",
      partitions: [
        {
          id: "p3",
          driveLetter: "E:\\",
          driveType: "Data",
          items: [
            {
              id: "i4",
              itemType: "File",
              name: "swapnil_migration_manifest",
              formatExtension: "csv",
              freeSpaceAvailable: "850.4 of 1000 GB",
              notes: "All corporate devices migration matrix spreadsheet",
            },
            {
              id: "i5",
              itemType: "Backup",
              name: "Outlook_Anjali_Backup_2024",
              formatExtension: "pst",
              freeSpaceAvailable: "420 of 1000 GB",
              notes: "Outlook email folder copy including attachments up to Q3 2024",
            },
          ],
        },
      ],
    },
    {
      srNo: 3,
      allottedTo: "John Doe",
      brand: "Crucial",
      type: "NVMe",
      capacity: "1 TB",
      serialNo: "HDD-C4F56A",
      dateAllotted: "2025-02-14",
      informationUrl: "#subsheet-HDD-C4F56A",
      partitions: [
        {
          id: "p4",
          driveLetter: "C:\\",
          driveType: "System",
          items: [
            {
              id: "i6",
              itemType: "Folder",
              name: "Users/John/Documents",
              formatExtension: "Directories",
              freeSpaceAvailable: "50 of 400 GB",
              notes: "Local work documents & profiles",
            },
          ],
        },
      ],
    },
  ];
  res.json({ success: true, database: inMemoryDatabase });
});

// Fetch Connection Status
app.get("/api/inventory/status", async (req, res) => {
  if (sheetCredentials.useSampleDatabase) {
    return res.json({
      connected: true,
      spreadsheetName: "Local Simulation Database Sandbox",
      spreadsheetId: "LOCAL_SANDBOX",
      sheetsCount: inMemoryDatabase.length + 1,
    });
  }

  if (sheetCredentials.appsScriptUrl) {
    return res.json({
      connected: true,
      spreadsheetName: "Google Apps Script Web App",
      spreadsheetId: "APPS_SCRIPT_SYNC",
      sheetsCount: 7, // Simulated sheets tab count
    });
  }

  try {
    const token = await getAuthToken();
    if (!token) {
      return res.json({
        connected: false,
        error: "Google credentials are empty or incomplete.",
      });
    }

    // Fetch spreadsheet metadata to check connection
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetCredentials.spreadsheetId}`;
    const metaRes = await fetch(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!metaRes.ok) {
      throw new Error(`Sheets API responded with ${metaRes.status}`);
    }

    const meta = await metaRes.json();
    res.json({
      connected: true,
      spreadsheetName: meta.properties.title || "Linked Google Sheet",
      spreadsheetId: sheetCredentials.spreadsheetId,
      sheetsCount: meta.sheets?.length || 0,
    });
  } catch (err: any) {
    res.json({
      connected: false,
      error: err.message || "Failed to establish validation link with Google Sheets API.",
    });
  }
});

// Fetch inventory list ('Main' sheet) and populate details
app.get("/api/inventory", async (req, res) => {
  if (sheetCredentials.useSampleDatabase) {
    return res.json({
      success: true,
      source: "Sandbox Local Buffer",
      items: inMemoryDatabase,
    });
  }

  if (sheetCredentials.appsScriptUrl) {
    try {
      const response = await fetch(sheetCredentials.appsScriptUrl + "?action=read");
      if (!response.ok) throw new Error(`Apps Script responded with status code ${response.status}`);
      const result = await response.json();
      if (result.success) {
        let webAppItems = result.items || [];
        webAppItems = webAppItems.map((itm: any) => {
          if (itm.partitionsRaw) {
            itm.partitions = parseSubSheetPartitionLayout(itm.partitionsRaw);
          }
          return itm;
        });
        return res.json({
          success: true,
          source: "Google Apps Script Sync",
          items: webAppItems,
        });
      } else {
        throw new Error(result.error || "Apps Script failed to fetch database rows.");
      }
    } catch (err: any) {
      console.error("Apps Script Bridge Read failed:", err);
      return res.status(550).json({ error: "Google Apps Script Bridge failed: " + err.message });
    }
  }

  try {
    const token = await getAuthToken();
    if (!token) throw new Error("Unauthorized sheet configuration.");

    const ssId = sheetCredentials.spreadsheetId;

    // Fetch master list from 'Main' tab
    const mainValues = await fetchSheetValues(ssId, "Main!A:H", token);
    const rows: any[][] = mainValues.values || [];

    if (rows.length === 0) {
      // Create 'Main' sheet if empty & seed with header
      const seedHeaders = [
        "Sr. No.",
        "HDD allotted",
        "Brand",
        "Type",
        "GB/TB",
        "Serial no.",
        "Date",
        "Information",
      ];
      await writeSheetValues(ssId, "Main!A1:H1", [seedHeaders], token);
      return res.json({ success: true, items: [], source: "Google Sheets (Seeded New)" });
    }

    // Parse the values row-by-row
    const headers = rows[0].map((h) => h?.toString().trim());
    const inventoryItems: HDDInventoryItem[] = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[1] && !r[5]) continue; // Skip empty rows

      const srNo = parseInt(r[0]) || i;
      const allottedTo = r[1] || "Allotted";
      const brand = r[2] || "Generic";
      const type = r[3] || "SSD";
      const capacity = r[4] || "Unknown";
      const serialNo = (r[5] || "").toString().trim();
      const dateAllotted = r[6] || "";
      const informationUrl = r[7] || "";

      // Load sub-sheet metadata to fetch the layout mapping
      let partitions: PartitionBlock[] = [];
      try {
        if (serialNo) {
          const detailRes = await fetchSheetValues(ssId, `'${serialNo}'!A1:Z100`, token);
          const detailRows: any[][] = detailRes.values || [];
          partitions = parseSubSheetPartitionLayout(detailRows);
        }
      } catch (err) {
        // Sub-sheet might not exist yet or have different tab name, fail silently
        console.warn(`Could not read sheet values for tab: ${serialNo}`);
      }

      inventoryItems.push({
        srNo,
        allottedTo,
        brand,
        type,
        capacity,
        serialNo,
        dateAllotted,
        informationUrl,
        partitions,
      });
    }

    res.json({
      success: true,
      source: "Google Sheets Live",
      items: inventoryItems,
    });
  } catch (err: any) {
    console.error("Fetch API error:", err);
    res.status(500).json({ error: err.message || "Failed to fetch sheets data." });
  }
});

// Sub-sheets parser layout mapper
function parseSubSheetPartitionLayout(rows: any[][]): PartitionBlock[] {
  const partitions: PartitionBlock[] = [];
  if (rows.length < 5) return partitions;

  let currentBlock: PartitionBlock | null = null;
  let expectItemRows = false;

  for (let idx = 3; idx < rows.length; idx++) {
    const cells = rows[idx];
    if (!cells || cells.length === 0) {
      if (currentBlock) {
        partitions.push(currentBlock);
        currentBlock = null;
        expectItemRows = false;
      }
      continue;
    }

    const firstCell = (cells[0] || "").toString().trim();

    // Check for standard partition section header: e.g. "C:\\ — System Drive" or "[Letter] — [Type] Drive"
    if (firstCell.includes(" — ")) {
      if (currentBlock) {
        partitions.push(currentBlock);
      }
      const parts = firstCell.split(" — ");
      const driveLetter = parts[0]?.trim() || "C:\\";
      const rest = parts[1]?.trim() || "";
      const isSystem = rest.toLowerCase().includes("system");
      const isBackup = rest.toLowerCase().includes("backup");
      const driveType = isSystem ? "System" : isBackup ? "Backup" : "Data";

      currentBlock = {
        id: `p-${idx}`,
        driveLetter,
        driveType,
        items: [],
      };
      expectItemRows = false;
      continue;
    }

    // Check for standard sub-headers row: [#', 'Item Type', 'Name / Description', 'Format / Extension', 'Free Space Available', 'Notes']
    if (firstCell === "#" || firstCell === "Sr") {
      expectItemRows = true;
      continue;
    }

    if (expectItemRows && currentBlock) {
      // Check if it's items limit/empty row
      if (firstCell && !isNaN(parseInt(firstCell))) {
        const itemType = cells[1] || "File";
        const name = cells[2] || "";
        const formatExtension = cells[3] || "";
        const freeSpaceAvailable = cells[4] || "";
        const notes = cells[5] || "";

        currentBlock.items.push({
          id: `item-${idx}`,
          itemType: itemType as any,
          name,
          formatExtension,
          freeSpaceAvailable,
          notes,
        });
      }
    }
  }

  if (currentBlock) {
    partitions.push(currentBlock);
  }

  return partitions;
}

// Smart HDD Intake Endpoint
app.post("/api/inventory/intake", async (req, res) => {
  const { allottedTo, brand, type, capacity, serialNo, dateAllotted, partitions, requestUser } = req.body;

  if (!serialNo) {
    return res.status(400).json({ error: "HDD Serial Number is required." });
  }

  const cleanSerial = serialNo.replace(/\s+/g, "-").toUpperCase();

  // 1. Google Apps Script pathway
  if (sheetCredentials.appsScriptUrl) {
    try {
      const response = await fetch(sheetCredentials.appsScriptUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "intake",
          allottedTo,
          brand,
          type,
          capacity,
          serialNo: cleanSerial,
          dateAllotted,
          partitions,
        }),
      });
      const result = await response.json();
      if (result.success) {
        await addAuditLog("Add HDD", `Added brand new HDD allotted to ${allottedTo} [Serial: ${cleanSerial}] (${capacity})`, requestUser);
        return res.json({
          success: true,
          message: "Ingested via Google Apps Script Web App successfully!",
          sheetLink: result.sheetLink || "#",
        });
      } else {
        throw new Error(result.error || "Web App script failed to record intake.");
      }
    } catch (err: any) {
      console.error("Apps Script Bridge Intake failed:", err);
      return res.status(500).json({ error: "Google Apps Script Ingestion failed: " + err.message });
    }
  }

  // 2. Sandbox simulation pathway
  if (sheetCredentials.useSampleDatabase) {
    const nextSr = inMemoryDatabase.length > 0 ? Math.max(...inMemoryDatabase.map((h) => h.srNo)) + 1 : 1;
    const mockItem: HDDInventoryItem = {
      srNo: nextSr,
      allottedTo: allottedTo || "Allotted",
      brand: brand || "Brand",
      type: type || "SSD",
      capacity: capacity || "512 GB",
      serialNo: cleanSerial,
      dateAllotted: dateAllotted || new Date().toISOString().split("T")[0],
      informationUrl: `#subsheet-${cleanSerial}`,
      partitions: partitions || [],
    };

    inMemoryDatabase.push(mockItem);
    await addAuditLog("Add HDD", `Added brand new Simulated HDD allotted to ${allottedTo} [Serial: ${cleanSerial}] (${capacity})`, requestUser);
    return res.json({
      success: true,
      message: "Successfully added to local simulation environment.",
      item: mockItem,
    });
  }

  // 3. SECURE GCP Service Account REST Sheets API pathway
  try {
    const token = await getAuthToken();
    if (!token) throw new Error("Google API connection token missing.");
    const ssId = sheetCredentials.spreadsheetId;

    // Create a brand-new sub-sheet tab in the spreadsheet
    console.log(`Creating tab for: ${cleanSerial}`);
    const addSheetBody = {
      requests: [
        {
          addSheet: {
            properties: {
              title: cleanSerial,
              gridProperties: {
                rowCount: 100,
                columnCount: 15,
              },
            },
          },
        },
      ],
    };

    let newSheetId = "";
    try {
      const sheetCreated = await updateSpreadsheet(ssId, addSheetBody, token);
      newSheetId = sheetCreated.replies[0].addSheet.properties.sheetId.toString();
    } catch (err: any) {
      if (err.message.includes("already exists")) {
        // Get sheet ID of existing tab
        const metaRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${ssId}`, {
          headers: { Authorization: `Bearer ${token}` },
        });
        const meta = await metaRes.json();
        const existingTab = meta.sheets?.find((s: any) => s.properties.title === cleanSerial);
        newSheetId = existingTab?.properties?.sheetId?.toString() || "";
      } else {
        throw err;
      }
    }

    // Format the Sub-Sheet data matching Section 2 Guidelines
    const detailValues: any[][] = [];
    detailValues.push(["HDD Data"]);
    detailValues.push(["User Name", "Brand", "Type", "Capacity", "Serial No.", "Date Alloted"]);
    detailValues.push([allottedTo, brand, type, capacity, cleanSerial, dateAllotted]);
    detailValues.push([]); // Row 4 space

    // Add partition blocks
    if (partitions && partitions.length > 0) {
      partitions.forEach((part: PartitionBlock) => {
        detailValues.push([`${part.driveLetter} — ${part.driveType} Drive`]);
        detailValues.push(["#", "Item Type", "Name / Description", "Format / Extension", "Free Space Available", "Notes"]);
        
        if (part.items && part.items.length > 0) {
          part.items.forEach((item: PartitionItem, idx: number) => {
            detailValues.push([
              idx + 1,
              item.itemType,
              item.name,
              item.formatExtension,
              item.freeSpaceAvailable,
              item.notes,
            ]);
          });
        } else {
          detailValues.push(["-", "Empty", "No Items Logged", "-", "N/A", "N/A"]);
        }
        detailValues.push([]); // spacer row after tables
      });
    }

    // Write content to sub-sheet
    await writeSheetValues(ssId, `'${cleanSerial}'!A1`, detailValues, token);

    // Append summary log entry into the 'Main' roster table
    let rowCount = 1;
    try {
      const getRes = await fetchSheetValues(ssId, "Main!A:A", token);
      rowCount = getRes.values?.length || 1;
    } catch (e) {
      // ignores error
    }

    // Formulate direct hyperlink url format
    const sheetHyperlink = `=HYPERLINK("https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=${newSheetId}", "View Tab: ${cleanSerial}")`;
    const mainRow = [
      rowCount,
      allottedTo,
      brand,
      type,
      capacity,
      cleanSerial,
      dateAllotted,
      sheetHyperlink,
    ];

    // Append beautiful values to next available row of 'Main'
    const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${ssId}/values/Main!A1:H1:append?valueInputOption=USER_ENTERED`;
    await fetch(appendUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: [mainRow] }),
    });

    await addAuditLog("Add HDD", `Added brand new Live Sheets HDD allotted to ${allottedTo} [Serial: ${cleanSerial}] (${capacity})`, requestUser);

    res.json({
      success: true,
      message: `HDD ${cleanSerial} ingested successfully! Formatted sub-sheet tab and synced main log.`,
      sheetLink: `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=${newSheetId}`,
    });
  } catch (err: any) {
    console.error("Core intake execution pipeline failed:", err);
    res.status(500).json({ error: err.message || "Integrative intake operation failed." });
  }
});

// Delete HDD record and wipe tab
app.post("/api/inventory/delete", async (req, res) => {
  const { serialNo, allottedTo, requestUser } = req.body;

  if (!serialNo) {
    return res.status(400).json({ error: "Serial number is required to perform deletion." });
  }

  const cleanSerial = serialNo.trim().toUpperCase();

  // 1. Apps Script pathway
  if (sheetCredentials.appsScriptUrl) {
    try {
      const response = await fetch(sheetCredentials.appsScriptUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "delete",
          serialNo: cleanSerial,
        }),
      });
      const result = await response.json();
      if (result.success) {
        await addAuditLog("Delete HDD", `Wiped HDD database asset ${cleanSerial} allotted to ${allottedTo || "N/A"}`, requestUser);
        return res.json({ success: true, message: `Wiped HDD ${cleanSerial} through Apps Script.` });
      } else {
        throw new Error(result.error || "Web App script failed to delete.");
      }
    } catch (err: any) {
      console.error("Apps Script Delete failed:", err);
      return res.status(500).json({ error: "Apps Script deletion failed: " + err.message });
    }
  }

  // 2. Sandbox Simulation Mode
  if (sheetCredentials.useSampleDatabase) {
    inMemoryDatabase = inMemoryDatabase.filter((h) => h.serialNo.toUpperCase() !== cleanSerial);
    await addAuditLog("Delete HDD", `Wiped HDD database asset ${cleanSerial} allotted to ${allottedTo || "N/A"} (Sandbox Simulation)`, requestUser);
    return res.json({ success: true, message: `Wiped Simulated HDD ${cleanSerial}.` });
  }

  // 3. GCP REST Sheets API Mode
  try {
    const token = await getAuthToken();
    if (!token) throw new Error("GCP Sheet connection unauthorized.");
    const ssId = sheetCredentials.spreadsheetId;

    // A. Fetch meta data & spreadsheets
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${ssId}`;
    const metaRes = await fetch(metaUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!metaRes.ok) throw new Error("Failed connecting sheets to execute wipe.");
    const meta = await metaRes.json();
    const mainSheet = meta.sheets?.find((s: any) => s.properties.title === "Main");
    const mainSheetId = mainSheet?.properties?.sheetId;

    if (mainSheetId === undefined) {
      throw new Error("Could not find Main roster worksheet in this Google Sheet.");
    }

    // B. Find row index of serial
    const mainValues = await fetchSheetValues(ssId, "Main!A:F", token);
    const rows: any[][] = mainValues.values || [];
    let targetRowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][5] || "").toString().trim().toUpperCase() === cleanSerial) {
        targetRowIndex = i;
        break;
      }
    }

    const requests: any[] = [];

    // C. Add deleteDimension requests for row
    if (targetRowIndex !== -1) {
      requests.push({
        deleteDimension: {
          range: {
            sheetId: mainSheetId,
            dimension: "ROWS",
            startIndex: targetRowIndex,
            endIndex: targetRowIndex + 1,
          },
        },
      });
    }

    // D. Check and add deleteSheet request for serial tab
    const serialTab = meta.sheets?.find((s: any) => s.properties.title.toUpperCase() === cleanSerial);
    if (serialTab) {
      requests.push({
        deleteSheet: {
          sheetId: serialTab.properties.sheetId,
        },
      });
    }

    if (requests.length > 0) {
      await updateSpreadsheet(ssId, { requests }, token);
    }

    await addAuditLog("Delete HDD", `Wiped HDD database asset ${cleanSerial} allotted to ${allottedTo || "N/A"} (Google Sheets Live)`, requestUser);
    res.json({ success: true, message: `Wiped master log row and sub-sheet tab for serial ${cleanSerial}.` });
  } catch (err: any) {
    console.error("Sheets delete api error:", err);
    res.status(500).json({ error: err.message || "Failed deleting spreadsheet records." });
  }
});

// GET Audit Logs Activity Feed
app.get("/api/audit-logs", async (req, res) => {
  // If livesheet connected and not using apps script
  if (!sheetCredentials.useSampleDatabase && !sheetCredentials.appsScriptUrl && sheetCredentials.spreadsheetId) {
    try {
      const token = await getAuthToken();
      if (token) {
        await ensureAuditLogSheetExists(sheetCredentials.spreadsheetId, token);
        const sheetRes = await fetchSheetValues(sheetCredentials.spreadsheetId, "AuditLog!A2:D500", token);
        const rows = sheetRes.values || [];
        const logsParsed: AuditLogItem[] = rows.map((r: any) => ({
          timestamp: r[0] || "",
          user: r[1] || "",
          action: r[2] || "",
          details: r[3] || "",
        })).reverse(); // newest first
        return res.json({ success: true, logs: logsParsed.length > 0 ? logsParsed : inMemoryAuditLogs });
      }
    } catch (err) {
      console.warn("Could not retrieve audit logs from Sheets, returning local session logs:", err);
    }
  }

  // If live Google Apps Script web app is connected
  if (!sheetCredentials.useSampleDatabase && sheetCredentials.appsScriptUrl) {
    try {
      const response = await fetch(sheetCredentials.appsScriptUrl + "?action=get_audit");
      if (response.ok) {
        const result = await response.json();
        if (result.success && result.logs) {
          return res.json({ success: true, logs: result.logs });
        }
      }
    } catch (err) {
      console.warn("Apps Script audit fetch error, returning local:", err);
    }
  }

  res.json({ success: true, logs: inMemoryAuditLogs });
});

// Gemini-AI Route: Section 3B. Smart terminal parser with Gemini model 'gemini-3.5-flash'
app.post("/api/ai/parse-text", async (req, res) => {
  const { rawText } = req.body;

  if (!rawText || rawText.trim().length === 0) {
    return res.status(400).json({ error: "Paste raw logs text is required to run the AI engine." });
  }

  const systemInstruction = `
You are an expert systems engineering metadata parser. The user will provide a rough, messy, or terminal text dump detailing the physical HDD and its directory/partition listings.
Your task is to parse this messy log into structured, highly accurate JSON.

Analyze the context for:
1. "allottedTo": Employee/Eng name (if found, otherwise default to "Corporate User")
2. "brand": Brand name (e.g. Western Digital, Seagate, Intel, Crucial, SanDisk, Samsung, Kingstown, etc.)
3. "type": HDD drive type, strictly choose from: "SSD", "HDD", "NVMe", "External", "Internal"
4. "capacity": Total storage size (e.g. "1 TB", "512 GB", "2 TB", "256 GB")
5. "serialNo": Unique Hard disk serial key. E.g. HDD-W2A89S, HDD-Z3980X (look for patterns like Serial: XXX, Serial No:, Model S/N, S/N, HDD-XXXX)
6. "dateAllotted": An ISO date format "YYYY-MM-DD". If not found, use current date "2026-06-02".
7. "partitions": Dynamic list of mapped storage parts/drives. E.g. C:\\, D:\\, E:\\ (Choose driveLetters cleanly).
   For EACH partition, detect the "driveType" ("System", "Data", "Backup") and extract listed "items" within it:
   Each item contains:
   - "itemType": Folder, File, Outlook PST, Backup, or Other. (If files like .pst or outlook emails are found, mark as "Outlook PST". If .zip/.tar, mark as "Backup". If folders, mark as "Folder").
   - "name": Clean item file folder name.
   - "formatExtension": File format/extension e.g. "pst", "zip", "sys", "exe", "Directories"
   - "freeSpaceAvailable": Remaining capacity or file sizing inside index, strictly formulated like "85.2 of 146 GB" or size metrics "42 GB".
   - "notes": Short summary of item purpose detected from comments.

Always return a valid JSON object matching this schema. Avoid markdown text wrappers other than the json layout. Keep schema naming exact.
`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: rawText,
      config: {
        systemInstruction,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          required: ["allottedTo", "brand", "type", "capacity", "serialNo", "dateAllotted", "partitions"],
          properties: {
            allottedTo: { type: Type.STRING },
            brand: { type: Type.STRING },
            type: { type: Type.STRING, description: 'Must be "SSD", "HDD", "NVMe", "External", or "Internal"' },
            capacity: { type: Type.STRING },
            serialNo: { type: Type.STRING },
            dateAllotted: { type: Type.STRING },
            partitions: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                required: ["driveLetter", "driveType", "items"],
                properties: {
                  driveLetter: { type: Type.STRING },
                  driveType: { type: Type.STRING, description: 'Must be "System", "Data", "Backup", or "Other"' },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      required: ["itemType", "name", "formatExtension", "freeSpaceAvailable", "notes"],
                      properties: {
                        itemType: { type: Type.STRING, description: 'Must be "Folder", "File", "Outlook PST", "Backup", or "Other"' },
                        name: { type: Type.STRING },
                        formatExtension: { type: Type.STRING },
                        freeSpaceAvailable: { type: Type.STRING },
                        notes: { type: Type.STRING },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    });

    const parsedJson = JSON.parse(response.text || "{}");
    res.json({ success: true, data: parsedJson });
  } catch (err: any) {
    console.error("AI parse execution error:", err);
    res.status(500).json({ error: err.message || "Failed to trigger AI parsing pipelines." });
  }
});

// Gemini-AI Route: Tab 3. AI Global Cross-Sheet Search
app.post("/api/ai/search", async (req, res) => {
  const { query, activeDatabase } = req.body;

  if (!query) {
    return res.status(400).json({ error: "Search query is required." });
  }

  // Fallback database source selector
  let dbToSearch: HDDInventoryItem[] = inMemoryDatabase;

  if (!sheetCredentials.useSampleDatabase) {
    // If using live Sheets, let client transfer the fetched inventory items representation
    if (activeDatabase && Array.isArray(activeDatabase)) {
      dbToSearch = activeDatabase;
    }
  }

  // Model context formatting
  let matrixContext = "DATABASE CONTEXT FOR CORPORATE HARD HARD DRIVES:\n";
  dbToSearch.forEach((hdd) => {
    matrixContext += `----------------------------------------\n`;
    matrixContext += `SERIAL_NO: ${hdd.serialNo}\n`;
    matrixContext += `ALLOTTED_TO: ${hdd.allottedTo}\n`;
    matrixContext += `DRIVE_BRAND_INFO: ${hdd.brand} (${hdd.type}) - Capacity: ${hdd.capacity}\n`;
    matrixContext += `DATE_ALLOTTED: ${hdd.dateAllotted}\n`;
    matrixContext += `PARTITIONS_LAYOUT:\n`;

    if (hdd.partitions && hdd.partitions.length > 0) {
      hdd.partitions.forEach((part) => {
        matrixContext += `  - Partition [${part.driveLetter}] (${part.driveType} Drive):\n`;
        if (part.items && part.items.length > 0) {
          part.items.forEach((item) => {
            matrixContext += `    * ItemType: ${item.itemType} | Name: ${item.name} | Ext: ${item.formatExtension} | Size/FreeSpace: ${item.freeSpaceAvailable} | Notes: ${item.notes}\n`;
          });
        } else {
          matrixContext += `    * No specific directory items logged on this drive.\n`;
        }
      });
    } else {
      matrixContext += `  No partitions logged yet.\n`;
    }
  });

  const promptText = `
Given the Corporate hard disk inventory spreadsheet context provided below, please answer the following user query:
"${query}"

Sheet Database Context:
${matrixContext}

Help the user trace where files or folders are located exactly. Include the owning employee, hard disk serial, specific partition, filename/item, and any relevant notes.
Please return your response in a beautiful JSON format containing two properties:
1. "markdownAnswer": A clean, scannable Markdown description explaining the tracking mapping. Be detailed and concise.
2. "isMatchFound": A boolean indicating if any drive matches the criteria.
3. "matchedItems": An array of matched metadata structures:
   [ { "owner": "Swapnil Shinde", "serial": "HDD-W2A89S", "driveLetter": "C:\\", "fileName": "corporate_archive_2024", "fileNotes": "mid-2024 outlook data" } ]
`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: promptText,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          required: ["markdownAnswer", "isMatchFound", "matchedItems"],
          properties: {
            markdownAnswer: { type: Type.STRING },
            isMatchFound: { type: Type.BOOLEAN },
            matchedItems: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                required: ["owner", "serial", "driveLetter", "fileName", "fileNotes"],
                properties: {
                  owner: { type: Type.STRING },
                  serial: { type: Type.STRING },
                  driveLetter: { type: Type.STRING },
                  fileName: { type: Type.STRING },
                  fileNotes: { type: Type.STRING },
                },
              },
            },
          },
        },
      },
    });

    res.json(JSON.parse(response.text || "{}"));
  } catch (err: any) {
    console.error("AI Cross sheet search prompt error:", err);
    res.status(500).json({ error: err.message || "Failed to query AI index." });
  }
});

// Configure Vite integration for dev server or serve production dist
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Smart HDD Inventory backend running on http://0.0.0.0:${PORT}`);
    console.log(`HMR status and watchers decoupled for platform performance`);
  });
}

startServer();
