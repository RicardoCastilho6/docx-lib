module.exports = {
  DOCUMENT_XML: "word/document.xml",
  RELS_FILE: "word/_rels/document.xml.rels",
  MEDIA_FOLDER: "word/media",

  MATCH_BODY: /<w:body>(.*?)<\/w:body>/s,
  MATCH_RELATIONSHIP: /<Relationship[\s\S]*?\/>/g,
  REPLACE_BODY: /<w:body>.*?<\/w:body>/s,

  STYLE_FILES: [
    "word/styles.xml",
    "word/theme/theme1.xml",
    "word/settings.xml",
  ],

  COMPRESSION_LEVEL: 9,
  ENCODING: "utf-8",
  RELS_END_TAG: "</Relationships>",

  TEMP_DIR_MERGE: "../../tmp/docx/temp_merge",
  TEMP_DIR_OUTPUT: "../../tmp/docx/output_merge",

  NS_CHECK: {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml"
  },

  ERROR_INVALID_FOLDER_LIST: "Lista de pastas inválida ou vazia.",
  ERROR_FOLDER_NOT_FOUND: "Pasta não encontrada: {folder}",
  SUCCESS_MERGE: "DOCX mesclado criado: {outputFilePath}",

  CONTENT_TYPES_FILE: "[Content_Types].xml",
  XML: 'application/xml',
  ENCODE_DEFAULT: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  W_NS: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  MERGE_BODY: ({ mergedBody }) => `<w:body>${mergedBody}</w:body>`,
};
