const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const docx = require('docx');
const fs = require('fs');

const app = express();

app.use(cors());
app.use(bodyParser.json());

const { Document, Packer, Paragraph, TextRun } = docx;

app.get('/', (req, res) => {
    res.send('Hello World!');
});

app.get('/generate-doc', (req, res) => {
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun("Hello World!"),
                            new TextRun({
                                text: "Bold text",
                                bold: true,
                            }),
                            new TextRun({
                                text: "\tTabbed text",
                            }),
                        ],
                    }),
                ],
            },
        ],
    });

    // Used for creating the Word Buffer
    Packer.toBuffer(doc).then(buffer => {
        fs.writeFileSync("Document.docx", buffer);
        res.download('Document.docx', 'GeneratedDocument.docx', (err) => {
            if (err) {
                res.status(500).send({
                    message: "Could not download the file. " + err,
                });
            }
            // Delete the file after sending it to the client
            fs.unlinkSync('Document.docx');
        });
    });
});

const PORT = 8080;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
