const fs = require('fs');
const path = require("path");
const Papa = require('papaparse');
const Docxtemplater = require('docxtemplater');
const PizZip = require("pizzip");

const readLetterTemplateFile = () => {
  return fs.readFileSync(
    path.resolve(__dirname, "letter-template.docx"),
    "binary"
  );
}

const readEnvelopeTemplateFile = () => {
  return fs.readFileSync(
    path.resolve(__dirname, "envelope-template.docx"),
    "binary"
  );
}

const csvData = fs.readFileSync(path.resolve(__dirname, "seller-data.csv"), 'utf8');
const csvResults = Papa.parse(csvData, {
  header: true,
  dynamicTyping: true,
  complete: () => {
    console.log("Finished reading CSV data");
  }
});

const mailingAddresses = []
const ownerNames = csvResults.data.filter(row => !!row.Owner).map((row) => {
  const { Owner: owner, ['Mailing Address 1']: mailingAddress1, ['Mailing Address 3']: mailingAddress3 } = row
  if(owner.length === 1 || owner.toUpperCase().includes('LLC')){
    mailingAddresses.push({owner, mailingAddress1, mailingAddress3})
    return owner
  }
  const ownerNamesSplit = owner.split(' ')
  const ownerName = ownerNamesSplit.map(x => {
    if( !x.length || /^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/gm.test(x)) {
      return x
    }
    else {
      return x.charAt(0).toUpperCase() + x.slice(1).toLowerCase();
    }
  }).join(' ')
  mailingAddresses.push({owner: ownerName, mailingAddress1, mailingAddress3})
  return ownerName
});

const letterTemplate = readLetterTemplateFile();
const letterDoc = new Docxtemplater(new PizZip(letterTemplate), {
    paragraphLoop: true,
    linebreaks: true,
});

const envelopeTemplate = readEnvelopeTemplateFile();
const envelopeDoc = new Docxtemplater(new PizZip(envelopeTemplate), {
  paragraphLoop: true,
  linebreaks: true,
})

letterDoc.render({
  loop: ownerNames.map((x) => {
    return { owner: x }
  }),
});
envelopeDoc.render({
  loop: mailingAddresses.map((x) => {
    return {
      ...x
    }
  })
})

const letterBuffer = letterDoc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});
const envelopeBuffer = envelopeDoc.getZip().generate({
  type: "nodebuffer",
  compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "generated-seller-letters.docx"), letterBuffer);
fs.writeFileSync(path.resolve(__dirname, "generated-envelopes.docx"), envelopeBuffer);