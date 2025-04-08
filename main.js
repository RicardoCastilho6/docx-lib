const { mergeDocxFolders } = require("./constants/services/docx-service");

(async () => {
  try {
    const files = [
      "doc1.docx",
      "doc2.docx",
      "doc3.docx",
      "doc4.docx",
      "doc5.docx",
      "doc6.docx",
      "doc7.docx",
      "doc8.docx",
      "doc9.docx",
      "doc10.docx",
    ];

    const outputFile = "merged.docx";

    await mergeDocxFolders(files, outputFile);

    console.log("DOCXs mesclados com sucesso!");
  } catch (error) {
    console.error("Erro ao mesclar DOCXs:", error);
  }
})();
