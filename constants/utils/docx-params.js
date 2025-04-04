module.exports = {
    DOCUMENT_XML: "word/document.xml",
    RELS_FILE: "word/_rels/document.xml.rels",
    MEDIA_FOLDER: "word/media",
    
    MATCH_BODY: /<w:body>(.*?)<\/w:body>/s,
    MATCH_RELATIONSHIP: /<Relationship[\s\S]*?\/>/g,
    REPLACE_BODY: /<w:body>.*?<\/w:body>/s,

    STYLE_FILES: ["word/styles.xml", "word/theme/theme1.xml", "word/settings.xml"],

    COMPRESSION_LEVEL: 9,
    ENCODING: "utf-8",
    RELS_END_TAG: "</Relationships>",

    TEMP_DIR_MERGE: "../../tmp/docx/temp_merge",
    TEMP_DIR_OUTPUT: "../../tmp/docx/output_merge",

    ERROR_INVALID_FOLDER_LIST: "Lista de pastas inválida ou vazia.",
    ERROR_FOLDER_NOT_FOUND: "Pasta não encontrada: {folder}",
    SUCCESS_MERGE: "DOCX mesclado criado: {outputFilePath}",

    MERGE_BODY: ({ mergedBody }) => `<w:body>${mergedBody}</w:body>`
};
