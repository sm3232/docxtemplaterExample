// Most of this is taken directly from docxtemplater docs.
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const content = fs.readFileSync(
    path.resolve(__dirname, "input.docx"),
    "binary"
);
const zip = new PizZip(content);
const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});
let obj = JSON.parse(
`{
    "first_name": "John",
    "last_name": "Smith",
    "address": "100 Main St",
    "city": "Boston",
    "state": "MA",
    "education": [
        { "date": "1994", "degree": "BS", "institution":"University of Chicago",
        "subject":"Business"}
    ],
    "work_history": [
        { "from": "1994", "to": "1997", "company":"Acme", "location":"New York, NY", "title":"Intern" },
        { "from": "1997", "to": "2000", "company":"Ajax", "location":"Los Angeles, CA", "title":"Jr Associate" },
        { "from": "2000", "to": "2003", "company":"Athena", "location":"Chicago, IL", "title":"Associate" }
    ]
}`
);
doc.render(obj);
const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
