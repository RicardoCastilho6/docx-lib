const { mergeDocxFolders } = require("./constants/services/docx-service");

(async () => {
    try {
        const folders = ["./tmp/xml/arquivo-um", "./tmp/xml/arquivo-dois", "./tmp/xml/arquivo-tres"];
        
        const outputFile = "merged.docx";

        await mergeDocxFolders(folders, outputFile);

        console.log("DOCXs mesclados com sucesso!");
    } catch (error) {
        console.error("Erro ao mesclar DOCXs:", error);
    }
})();
