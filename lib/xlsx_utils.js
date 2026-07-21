/**
 * Cross-platform XLSX pack/unpack and cell helpers.
 * Uses system zip/unzip (macOS/Linux) or PowerShell (Windows).
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Set or replace a cell value in sheet XML, preserving style when present.
 */
function setCellValue(sheetXml, ref, value, { asString = false } = {}) {
  const cellRe = new RegExp(`<c r="${ref}"([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`);
  const m = sheetXml.match(cellRe);

  if (value === null || value === undefined || value === "") {
    if (!m) return sheetXml;
    let attrs = m[1] || "";
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    return sheetXml.replace(cellRe, `<c r="${ref}"${attrs}/>`);
  }

  const writeAsString = asString || typeof value === "string";

  if (m) {
    let attrs = m[1] || "";
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    if (writeAsString) {
      return sheetXml.replace(
        cellRe,
        `<c r="${ref}"${attrs} t="inlineStr"><is><t>${xmlEscape(
          String(value)
        )}</t></is></c>`
      );
    }
    return sheetXml.replace(
      cellRe,
      `<c r="${ref}"${attrs}><v>${value}</v></c>`
    );
  }

  const rowNum = ref.replace(/^[A-Z]+/, "");
  const rowRe = new RegExp(`(<row r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
  if (!sheetXml.match(rowRe)) {
    throw new Error(`Row ${rowNum} not found for cell ${ref}`);
  }
  const newCell = writeAsString
    ? `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(
        String(value)
      )}</t></is></c>`
    : `<c r="${ref}"><v>${value}</v></c>`;
  return sheetXml.replace(rowRe, `$1$2${newCell}$3`);
}

function hasCmd(cmd) {
  try {
    execSync(
      process.platform === "win32" ? `where ${cmd}` : `command -v ${cmd}`,
      { stdio: "pipe" }
    );
    return true;
  } catch {
    return false;
  }
}

/**
 * Unpack an xlsx (zip) into a directory.
 */
function unpackXlsx(xlsxPath, destDir) {
  fs.mkdirSync(destDir, { recursive: true });
  const absXlsx = path.resolve(xlsxPath);
  const absDest = path.resolve(destDir);

  if (process.platform === "win32" && !hasCmd("unzip")) {
    execSync(
      `powershell -NoProfile -Command "Expand-Archive -LiteralPath '${absXlsx.replace(
        /'/g,
        "''"
      )}' -DestinationPath '${absDest.replace(/'/g, "''")}' -Force"`,
      { stdio: "pipe" }
    );
    return;
  }

  execSync(`unzip -q -o "${absXlsx}" -d "${absDest}"`, { stdio: "pipe" });
}

/**
 * Pack a directory into an xlsx (zip) file.
 */
function packXlsx(sourceDir, outputPath) {
  const absSrc = path.resolve(sourceDir);
  const absOut = path.resolve(outputPath);
  fs.mkdirSync(path.dirname(absOut), { recursive: true });
  if (fs.existsSync(absOut)) fs.unlinkSync(absOut);

  if (process.platform === "win32" && !hasCmd("zip")) {
    const ps1 = path.join(os.tmpdir(), `zip_xlsx_${Date.now()}.ps1`);
    fs.writeFileSync(
      ps1,
      [
        "Add-Type -AssemblyName System.IO.Compression.FileSystem",
        `if (Test-Path -LiteralPath '${absOut.replace(/'/g, "''")}') { Remove-Item -LiteralPath '${absOut.replace(/'/g, "''")}' -Force }`,
        `[System.IO.Compression.ZipFile]::CreateFromDirectory('${absSrc.replace(/'/g, "''")}', '${absOut.replace(/'/g, "''")}', [System.IO.Compression.CompressionLevel]::Optimal, $false)`,
      ].join("\r\n")
    );
    try {
      execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${ps1}"`, {
        stdio: "pipe",
      });
    } finally {
      try {
        fs.unlinkSync(ps1);
      } catch {
        /* ignore */
      }
    }
    return;
  }

  // zip contents of directory (not the directory itself)
  execSync(`cd "${absSrc}" && zip -qr "${absOut}" .`, { stdio: "pipe" });
}

/**
 * Copy template xlsx, yield unpacked dir, pack to output when done.
 * @returns {{ unpacked: string, tmp: string, finish: (outputPath: string) => void, cleanup: () => void }}
 */
function openTemplateCopy(templatePath) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ferc-xlsx-"));
  const zipCopy = path.join(tmp, "book.xlsx");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  unpackXlsx(zipCopy, unpacked);

  return {
    tmp,
    unpacked,
    finish(outputPath) {
      packXlsx(unpacked, outputPath);
    },
    cleanup() {
      fs.rmSync(tmp, { recursive: true, force: true });
    },
  };
}

/**
 * Optionally rename the first sheet in workbook.xml.
 */
function setSheetName(unpackedDir, sheetName) {
  if (!sheetName) return;
  const wbPath = path.join(unpackedDir, "xl", "workbook.xml");
  let wb = fs.readFileSync(wbPath, "utf8");
  if (!wb.includes(`name="${sheetName}"`)) {
    wb = wb.replace(
      /(<sheet\s+name=")[^"]+(")/,
      `$1${sheetName.replace(/"/g, "")}$2`
    );
    fs.writeFileSync(wbPath, wb, "utf8");
  }
}

module.exports = {
  xmlEscape,
  setCellValue,
  unpackXlsx,
  packXlsx,
  openTemplateCopy,
  setSheetName,
};
