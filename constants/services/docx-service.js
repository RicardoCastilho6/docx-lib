const fs = require("fs");
const path = require("path");
const archiver = require("archiver");
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("xmldom");
const DOCX = require("../utils/docx-params");
const { ensureNamespaces } = require("../utils/xml-utils");

const unzipDocx = async (docxPath, outputDir) => {
  try {
    const data = await fs.promises.readFile(docxPath);
    const zip = await JSZip.loadAsync(data);

    await Promise.all(
      Object.keys(zip.files).map(async (relativePath) => {
        const file = zip.files[relativePath];
        const fullPath = path.join(outputDir, relativePath);

        if (file.dir) {
          await fs.promises.mkdir(fullPath, { recursive: true });
        } else {
          const content = await file.async("nodebuffer");
          await fs.promises.mkdir(path.dirname(fullPath), { recursive: true });
          await fs.promises.writeFile(fullPath, content);
        }
      })
    );

    console.log(`${docxPath} descompactado em ${outputDir}`);
  } catch (err) {
    throw new Error(`Erro ao descompactar ${docxPath}: ${err.message}`);
  }
};

const mergeDocxFolders = async (files, outputFilePath) => {
  if (!Array.isArray(files) || files.length === 0) {
    throw new Error(DOCX.ERROR_INVALID_FOLDER_LIST);
  }

  const tempFolders = [];
  for (const file of files) {
    const tempDir = path.join(
      DOCX.TEMP_DIR_MERGE,
      path.basename(file, ".docx")
    );
    await unzipDocx(file, tempDir);
    tempFolders.push(tempDir);
  }

  const TEMP_DIR = path.join(DOCX.TEMP_DIR_MERGE, "merged");
  if (!fs.existsSync(TEMP_DIR)) {
    fs.mkdirSync(TEMP_DIR, { recursive: true });
  }

  tempFolders.forEach((folder, index) => {
    if (index === 0) {
      copyFolderRecursiveSync(folder, TEMP_DIR);
    } else {
      mergeDocxContent(folder, TEMP_DIR);
    }
  });

  const output = fs.createWriteStream(outputFilePath);
  const archive = archiver("zip", { zlib: { level: DOCX.COMPRESSION_LEVEL } });

  return new Promise((resolve, reject) => {
    output.on("close", () => {
      console.log(
        DOCX.SUCCESS_MERGE.replace("{outputFilePath}", outputFilePath)
      );
      fs.rmSync(DOCX.TEMP_DIR_MERGE, { recursive: true, force: true });
      resolve();
    });

    archive.on("error", (err) => reject(err));
    archive.pipe(output);
    archive.directory(TEMP_DIR, false);
    archive.finalize();
  });
};

const copyFolderRecursiveSync = (source, target) => {
  if (!fs.existsSync(target)) {
    fs.mkdirSync(target, { recursive: true });
  }

  fs.readdirSync(source).forEach((file) => {
    const srcPath = path.join(source, file);
    const destPath = path.join(target, file);

    if (fs.lstatSync(srcPath).isDirectory()) {
      copyFolderRecursiveSync(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  });
};

const mergeMedia = (sourceFolder, targetFolder, renamedMediaMap) => {
  const sourceMedia = path.join(sourceFolder, DOCX.MEDIA_FOLDER);
  const targetMedia = path.join(targetFolder, DOCX.MEDIA_FOLDER);

  if (!fs.existsSync(sourceMedia)) return;

  if (!fs.existsSync(targetMedia)) {
    fs.mkdirSync(targetMedia, { recursive: true });
  }

  const files = fs.readdirSync(sourceMedia);
  files.forEach((file) => {
    let destFile = path.join(targetMedia, file);
    let newFilename = file;
    let counter = 1;

    while (fs.existsSync(destFile)) {
      const ext = path.extname(file);
      const base = path.basename(file, ext);
      newFilename = `${base}_${counter}${ext}`;
      destFile = path.join(targetMedia, newFilename);
      counter++;
    }

    if (newFilename !== file) {
      renamedMediaMap.set(file, newFilename);
    }

    fs.copyFileSync(path.join(sourceMedia, file), destFile);
  });
};

const updateRelationships = (sourceFolder, targetFolder, renamedMediaMap) => {
  const sourceRelsPath = path.join(sourceFolder, DOCX.RELS_FILE);
  const targetRelsPath = path.join(targetFolder, DOCX.RELS_FILE);

  if (!fs.existsSync(sourceRelsPath)) return;

  const sourceRelsXml = fs.readFileSync(sourceRelsPath, DOCX.ENCODING);
  let targetRelsXml = fs.existsSync(targetRelsPath)
    ? fs.readFileSync(targetRelsPath, DOCX.ENCODING)
    : DOCX.ENCODE_DEFAULT;

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const sourceDoc = parser.parseFromString(sourceRelsXml, DOCX.XML);
  const targetDoc = parser.parseFromString(targetRelsXml, DOCX.XML);

  const sourceRels = sourceDoc.getElementsByTagName("Relationship");
  const targetRels = targetDoc.getElementsByTagName("Relationship");

  const existingIds = new Set();
  for (let i = 0; i < targetRels.length; i++) {
    existingIds.add(targetRels[i].getAttribute("Id"));
  }

  for (let i = 0; i < sourceRels.length; i++) {
    const sourceRel = sourceRels[i];
    const id = sourceRel.getAttribute("Id");
    if (!existingIds.has(id)) {
      const newRel = sourceRel.cloneNode(true);
      const targetAttr = newRel.getAttribute("Target");
      if (targetAttr && targetAttr.startsWith("../media/")) {
        const filename = path.basename(targetAttr);
        if (renamedMediaMap.has(filename)) {
          newRel.setAttribute(
            "Target",
            `../media/${renamedMediaMap.get(filename)}`
          );
        }
      }
      targetDoc.documentElement.appendChild(newRel);
      existingIds.add(id);
    }
  }

  const updatedXml = serializer.serializeToString(targetDoc);
  fs.writeFileSync(targetRelsPath, updatedXml, DOCX.ENCODING);
};

const mergeContentTypes = (sourceFolder, targetFolder) => {
  const sourcePath = path.join(sourceFolder, DOCX.CONTENT_TYPES_FILE);
  const targetPath = path.join(targetFolder, DOCX.CONTENT_TYPES_FILE);

  if (!fs.existsSync(sourcePath)) return;

  const sourceXml = fs.readFileSync(sourcePath, DOCX.ENCODING);
  const targetXml = fs.readFileSync(targetPath, DOCX.ENCODING);

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const sourceDoc = parser.parseFromString(sourceXml, DOCX.XML);
  const targetDoc = parser.parseFromString(targetXml, DOCX.XML);

  // Merge Defaults
  const sourceDefaults = sourceDoc.getElementsByTagName("Default");
  for (let i = 0; i < sourceDefaults.length; i++) {
    const def = sourceDefaults[i];
    const ext = def.getAttribute("Extension");
    const contentType = def.getAttribute("ContentType");
    const exists = Array.from(targetDoc.getElementsByTagName("Default")).some(
      (d) => d.getAttribute("Extension") === ext
    );
    if (!exists) {
      const imported = targetDoc.importNode(def, true);
      targetDoc.documentElement.insertBefore(
        imported,
        targetDoc.documentElement.firstChild
      );
    }
  }

  // Merge Overrides
  const sourceOverrides = sourceDoc.getElementsByTagName("Override");
  for (let i = 0; i < sourceOverrides.length; i++) {
    const ovr = sourceOverrides[i];
    const partName = ovr.getAttribute("PartName");
    const exists = Array.from(targetDoc.getElementsByTagName("Override")).some(
      (o) => o.getAttribute("PartName") === partName
    );
    if (!exists) {
      const imported = targetDoc.importNode(ovr, true);
      targetDoc.documentElement.appendChild(imported);
    }
  }

  fs.writeFileSync(
    targetPath,
    serializer.serializeToString(targetDoc),
    DOCX.ENCODING
  );
};

const mergeStylesAndSettings = (sourceFolder, targetFolder) => {
  DOCX.STYLE_FILES.forEach((file) => {
    const sourcePath = path.join(sourceFolder, file);
    const targetPath = path.join(targetFolder, file);

    if (fs.existsSync(sourcePath) && !fs.existsSync(targetPath)) {
      fs.copyFileSync(sourcePath, targetPath);
    }
  });
};

const mergeDocxContent = (sourceFolder, targetFolder) => {
  const docXmlPath1 = path.join(targetFolder, DOCX.DOCUMENT_XML);
  const docXmlPath2 = path.join(sourceFolder, DOCX.DOCUMENT_XML);

  if (fs.existsSync(docXmlPath1) && fs.existsSync(docXmlPath2)) {
    const content1 = fs.readFileSync(docXmlPath1, DOCX.ENCODING);
    const content2 = fs.readFileSync(docXmlPath2, DOCX.ENCODING);

    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    const doc1 = parser.parseFromString(content1, DOCX.XML);
    const doc2 = parser.parseFromString(content2, DOCX.XML);

    const documentElement = doc1.documentElement;
    
    ensureNamespaces(documentElement, DOCX.NS_CHECK)

    const body1 = doc1.getElementsByTagNameNS(DOCX.W_NS, "body")[0];
    const body2 = doc2.getElementsByTagNameNS(DOCX.W_NS, "body")[0];

    for (let i = 0; i < body2.childNodes.length; i++) {
      const node = body2.childNodes[i];
      body1.appendChild(node.cloneNode(true));
    }

    fs.writeFileSync(
      docXmlPath1,
      serializer.serializeToString(doc1),
      DOCX.ENCODING
    );

    const renamedMediaMap = new Map();
    mergeMedia(sourceFolder, targetFolder, renamedMediaMap);
    updateRelationships(sourceFolder, targetFolder, renamedMediaMap);
    mergeContentTypes(sourceFolder, targetFolder);
    mergeStylesAndSettings(sourceFolder, targetFolder);
  }
};

module.exports = { mergeDocxFolders };
