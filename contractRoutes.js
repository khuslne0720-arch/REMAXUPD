// contractRoutes.js
// npm install xlsx (нэг удаа суулгана)

const express = require("express");
const router = express.Router();
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const CONTRACTS_FILE = path.join(__dirname, "contracts.json");

// ─── Helper functions ───────────────────────────────────────────────

function readContracts() {
  if (!fs.existsSync(CONTRACTS_FILE)) return [];
  try {
    return JSON.parse(fs.readFileSync(CONTRACTS_FILE, "utf-8"));
  } catch {
    return [];
  }
}

function writeContracts(data) {
  fs.writeFileSync(CONTRACTS_FILE, JSON.stringify(data, null, 2), "utf-8");
}

// ─── 1. Гэрээ хадгалах ──────────────────────────────────────────────
// POST /contracts
// Body: { contractNumber, agent, startDate, endDate, propertyId,
//         listingType, area, address, owner, register }

router.post("/contracts", (req, res) => {
  const {
    contractNumber, agent, startDate, endDate,
    propertyId, listingType, area, address, owner, register,
  } = req.body;

  if (!contractNumber) {
    return res.status(400).json({ error: "Гэрээний дугаар заавал шаардлагатай." });
  }

  const contracts = readContracts();

  // Давхардал шалгах
  const exists = contracts.find((c) => c.contractNumber === contractNumber);
  if (exists) {
    return res.status(409).json({ error: "Энэ дугаартай гэрээ аль хэдийн байна." });
  }

  const newContract = {
    id: Date.now().toString(),
    contractNumber,
    agent,
    startDate,
    endDate,
    propertyId,
    listingType,
    area,
    address,
    owner,
    register,
    createdAt: new Date().toISOString(),
  };

  contracts.push(newContract);
  writeContracts(contracts);

  res.status(201).json({ message: "Гэрээ амжилттай хадгалагдлаа.", contract: newContract });
});

// ─── 2. Гэрээ preview (нэг гэрээ харах) ────────────────────────────
// GET /contracts/:id

router.get("/contracts/:id", (req, res) => {
  const contracts = readContracts();
  const contract = contracts.find((c) => c.id === req.params.id);

  if (!contract) {
    return res.status(404).json({ error: "Гэрээ олдсонгүй." });
  }

  res.json(contract);
});

// ─── 3. Гэрээ засах (edit) ──────────────────────────────────────────
// PUT /contracts/:id

router.put("/contracts/:id", (req, res) => {
  const contracts = readContracts();
  const index = contracts.findIndex((c) => c.id === req.params.id);

  if (index === -1) {
    return res.status(404).json({ error: "Гэрээ олдсонгүй." });
  }

  // Зөвхөн зөвшөөрөгдсөн талбаруудыг шинэчлэх
  const allowed = [
    "contractNumber", "agent", "startDate", "endDate",
    "propertyId", "listingType", "area", "address", "owner", "register",
  ];

  allowed.forEach((field) => {
    if (req.body[field] !== undefined) {
      contracts[index][field] = req.body[field];
    }
  });

  contracts[index].updatedAt = new Date().toISOString();
  writeContracts(contracts);

  res.json({ message: "Гэрээ амжилттай шинэчлэгдлаа.", contract: contracts[index] });
});

// ─── 4. Гэрээ устгах ────────────────────────────────────────────────
// DELETE /contracts/:id

router.delete("/contracts/:id", (req, res) => {
  const contracts = readContracts();
  const filtered = contracts.filter((c) => c.id !== req.params.id);

  if (filtered.length === contracts.length) {
    return res.status(404).json({ error: "Гэрээ олдсонгүй." });
  }

  writeContracts(filtered);
  res.json({ message: "Гэрээ устгагдлаа." });
});

// ─── 5. Excel export (ЗӨВХӨН ADMIN) ────────────────────────────────
// GET /admin/export-excel
// Header: x-admin-key: <таны нууц түлхүүр>

const ADMIN_KEY = process.env.ADMIN_KEY || "CHANGE_THIS_SECRET_KEY";

router.get("/admin/export-excel", (req, res) => {
  // Admin эрх шалгах
  const key = req.headers["x-admin-key"];
  if (key !== ADMIN_KEY) {
    return res.status(403).json({ error: "Хандах эрх байхгүй." });
  }

  const contracts = readContracts();

  if (contracts.length === 0) {
    return res.status(400).json({ error: "Гэрээний мэдээлэл хоосон байна." });
  }

  // Excel-д гаргах өгөгдөл (монгол гарчигтай)
  const rows = contracts.map((c) => ({
    "Гэрээний дугаар":          c.contractNumber ?? "",
    "Агент":                    c.agent ?? "",
    "Эхлэх":                    c.startDate ?? "",
    "Дуусах":                   c.endDate ?? "",
    "ҮХХ дугаар":               c.propertyId ?? "",
    "Листингийн төрөл":         c.listingType ?? "",
    "Талбайн хэмжээ":           c.area ?? "",
    "Листингийн байршил":       c.address ?? "",
    "ҮХХ эзэмшигчийн мэдээлэл": c.owner ?? "",
    "РД":                       c.register ?? "",
  }));

  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(rows);

  // Баганы өргөн тохируулах
  worksheet["!cols"] = [
    { wch: 18 }, // Гэрээний дугаар
    { wch: 16 }, // Агент
    { wch: 12 }, // Эхлэх
    { wch: 12 }, // Дуусах
    { wch: 14 }, // ҮХХ дугаар
    { wch: 18 }, // Листингийн төрөл
    { wch: 16 }, // Талбайн хэмжээ
    { wch: 30 }, // Листингийн байршил
    { wch: 28 }, // ҮХХ эзэмшигчийн мэдээлэл
    { wch: 12 }, // РД
  ];

  xlsx.utils.book_append_sheet(workbook, worksheet, "Гэрээнүүд");

  const buffer = xlsx.write(workbook, { type: "buffer", bookType: "xlsx" });

  const filename = `contracts_${new Date().toISOString().slice(0, 10)}.xlsx`;

  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});

module.exports = router;
