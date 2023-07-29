const Papa = require('papaparse');
const fs = require('fs');
const officegen = require('officegen');
const docx = officegen('docx');

docx.on('finalize', function() {
  console.log('Finished generating Word document');
});

docx.on('error', function(err) {
  console.log('Error generating Word document', err);
});

const csvData = fs.readFileSync('./seller-data.csv', 'utf8');
const results = Papa.parse(csvData, {
  header: true,
  dynamicTyping: true,
  complete: function() {
    console.log("Finished reading CSV data");
  }
});

const ownerNames = results.data.filter(row => !!row.Owner).map((row) => {
  const { Owner } = row
  if(Owner.length === 1 || Owner.toUpperCase().includes('LLC')){
    return Owner
  }
  const ownerNamesSplit = Owner.split(' ')
  return ownerNamesSplit.map(x => {
    if( !x.length || /^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/gm.test(x)) {
      return x
    }
    else {
      return x.charAt(0).toUpperCase() + x.slice(1).toLowerCase();
    }
  }).join(' ')
});

const standardOptions = {
  font_face: 'Source Sans Pro',
  font_size: 12,
}

const boldOptions = {
  ...standardOptions,
  bold: true,
}

const dataMapped = ownerNames.map((ownerName) => {
  return `${ownerName}, we first want to thank you for taking your time to read this letter. We are GHS Investment Group, a premier home buying service based out of Atlanta, Georgia. GHS specializes in helping homeowners with their real estate problems such as: unwanted rental property, foreclosure, vacant properties, inherited properties and many more. Our goal is to cultivate an offer that meets your needs as well as allows us to make a profit to stay in business.`;
});

dataMapped.forEach((entry) => {
  const pObj = docx.createP();
  pObj.addText(entry, standardOptions);
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText('Our offer is based on a few different factors including...', standardOptions)
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText('Comparable Sales.', boldOptions)
  pObj.addLineBreak();
  pObj.addText(`We use Fannie Mae's appraisal guidelines for determining comparable/similar properties. Fannie Mae's guidelines says to find 3 comparable (comp) properties that have `, standardOptions)
  pObj.addText(` SOLD within the last 6 months, similar condition`, boldOptions)
  pObj.addText(`, are `, standardOptions)
  pObj.addText('within 1 mile', boldOptions)
  pObj.addText(' of the subject property, the ', standardOptions)
  pObj.addText(`square footage is within 10%`, boldOptions)
  pObj.addText(' of the subject property and are ', standardOptions)
  pObj.addText('built within 10 years', boldOptions)
  pObj.addText(' of the subject property.', standardOptions)
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText(`Also keep in mind that these values are not what the sellers netted in their pocket. They also had to pay commissions (6%), closing costs (2%), holding costs (2%), and even make repairs for the buyer up to 15% of the home's value.`, standardOptions)
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText('Repairs.', boldOptions)
  pObj.addLineBreak();
  pObj.addText(`We also took into consideration the amount of repairs it would cost us to renovate your property and bring it to today's selling standards. Today's buyers EXPECT things like granite countertops, stainless steel appliances, open concept floor plans, etc. As well as bringing things up to today's building codes in regards to plumbing, electrical, etc. And like 99% of renovation as small as changing carpet to as big as building a brand new home, the renovation budget always seems to be more than anticipated.`, standardOptions);
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText('Our Offer.', boldOptions)
  pObj.addLineBreak();
  pObj.addText(`The last thing we factor is a profit margin of about 10-15%, depending upon the amount of work needed to be done & the length of time to complete the project.  So we carefully factor all these variables into crafting a fair offer. Also remember we're closing with CASH & buying your home in AS-IS condition. We'll pay all your closing costs. And there are absolutely NO commissions or fees you have to pay.`, standardOptions);
  pObj.addLineBreak();
  pObj.addLineBreak();
  pObj.addText(`You're getting a guarantee price, & closing date without any contingencies, so that you can start to plan for the future.`, standardOptions)
});

let out = fs.createWriteStream('generated-seller-letters.docx');

out.on('error', function(err) {
  console.log(err);
});

docx.generate(out);
