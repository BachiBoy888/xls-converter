import express from "express";
import cors from "cors";
import multer from "multer";
import XLSX from "xlsx";

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: (process.env.MAX_FILE_SIZE_MB ? +process.env.MAX_FILE_SIZE_MB : 10) * 1024 * 1024 }
});

// healthcheck
app.get("/healthz", (_req, res) => res.status(200).send("ok"));

// XLS/XLSX -> XLSX
app.post("/convert/xlsx", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Attach file field name 'file'." });

    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const outBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    const base = (req.file.originalname || "converted").replace(/\.[^.]+$/, "");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${base}.xlsx"`);
    return res.send(outBuffer);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Failed to convert file to XLSX." });
  }
});

// 404
app.use((_req, res) => res.status(404).json({ error: "Not found" }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`XLS converter listening on ${PORT}`));
