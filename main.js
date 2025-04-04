const { mergeDocxFolders } = require("./constants/services/docx-service");

(async () => {
    try {
        const files = ["1.docx", "2.docx", "3.docx"];
        
        const outputFile = "merged.docx";

        await mergeDocxFolders(files, outputFile);

        console.log("DOCXs mesclados com sucesso!");
    } catch (error) {
        console.error("Erro ao mesclar DOCXs:", error);
    }
})();
