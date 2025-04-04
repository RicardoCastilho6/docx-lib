const fs = require("fs");
const path = require("path");
const archiver = require("archiver");
const JSZip = require("jszip")
const DOCX = require("../utils/docx-params");

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

const mergeMedia = (sourceFolder, targetFolder) => {
  const sourceMedia = path.join(sourceFolder, DOCX.MEDIA_FOLDER);
  const targetMedia = path.join(targetFolder, DOCX.MEDIA_FOLDER);

  if (!fs.existsSync(sourceMedia)) return;

  if (!fs.existsSync(targetMedia)) {
    fs.mkdirSync(targetMedia, { recursive: true });
  }

  fs.readdirSync(sourceMedia).forEach((file) => {
    const srcFile = path.join(sourceMedia, file);
    const destFile = path.join(targetMedia, file);

    if (!fs.existsSync(destFile)) {
      fs.copyFileSync(srcFile, destFile);
    }
  });
};

const updateRelationships = (sourceFolder, targetFolder) => {
  const sourceRelsPath = path.join(sourceFolder, DOCX.RELS_FILE);
  const targetRelsPath = path.join(targetFolder, DOCX.RELS_FILE);

  if (!fs.existsSync(sourceRelsPath)) return;

  let sourceRels = fs.readFileSync(sourceRelsPath, DOCX.ENCODING);
  let targetRels = fs.readFileSync(targetRelsPath, DOCX.ENCODING);

  const relationships = sourceRels.match(DOCX.MATCH_RELATIONSHIP) || [];
  relationships.forEach((rel) => {
    if (!targetRels.includes(rel)) {
      targetRels = targetRels.replace(
        DOCX.RELS_END_TAG,
        `${rel}\n${DOCX.RELS_END_TAG}`
      );
    }
  });

  fs.writeFileSync(targetRelsPath, targetRels, DOCX.ENCODING);
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

    const body1 = content1.match(DOCX.MATCH_BODY)?.[1] || "";
    const body2 = content2.match(DOCX.MATCH_BODY)?.[1] || "";
    const mergedBody = body1 + body2;

    const mergedContent = content1.replace(
      DOCX.REPLACE_BODY,
      DOCX.MERGE_BODY({ mergedBody })
    );

    fs.writeFileSync(docXmlPath1, mergedContent, DOCX.ENCODING);

    mergeMedia(sourceFolder, targetFolder);
    updateRelationships(sourceFolder, targetFolder);
    mergeStylesAndSettings(sourceFolder, targetFolder);
  }
};

module.exports = { mergeDocxFolders };
