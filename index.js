const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
require("dotenv").config();

const fs = require("fs");
const path = require("path");
const express = require("express");
const app = express();
const port = process.env.PORT || 3000;

app.get("/", (req, res) => {
  // Load the docx file as binary content
  const content = fs.readFileSync(
    path.resolve(__dirname, "temp.docx"),
    "binary"
  );

  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
  doc.render({
    nome: "John",
    cpf: "123456789",
  });
  console.log("docx render");
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
  });

  // buf is a nodejs Buffer, you can either write it to a
  // file or res.send it with express for example.
  fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
  console.log("docx criado");
  const { convertWordFiles } = require("convert-multiple-files");

  const docToPdf = async (filePath, outputDir) =>
    await convertWordFiles(filePath, "pdf", outputDir);

  docToPdf(path.resolve(__dirname, "output.docx"), path.resolve(__dirname));
  console.log("pdf criado");
  return res.send("finish");
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
