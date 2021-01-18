const fs = require('fs');
const pdf = require('pdf-parse');

let dataBuffer = fs.readFileSync('Abstract_Book.pdf');
// Find next poster
function findNextItem(str, pos, item) {
  itemstr = ('00' + item).slice(-3);
  itemstr = 'P' + itemstr;
  return str.indexOf(itemstr, pos);
}

function splitTextIntoPosters(msg) {
  var posterArr = [];
  var index = 1;
  var pos = 0;
  while (pos >= 0) {
    posprev = pos;
    pos = findNextItem(msg, posprev, index);
    if (index > 1) {
      if (pos > 0) {
        posterArr.push(msg.substring(posprev, pos));
      } else {
        posterArr.push(msg.substring(posprev));
      }
    }
    index++;
  }
  return posterArr;
}

function splitText(str) {

  // remove all footnotes
  str = str.replace(/\n\d\n/g, '');
  // remove all hyphens
  str = str.replace(/\-\n/g, '');

	// JSON we parse into xlsx later
  let dataArr = [];

  // split the string into an array without delimiter removing
  let posterArr = splitTextIntoPosters(str);

  posterArr.forEach(section => {
    let obj = {
      name: [],
      affiliationName: '',
      personsLocation: '',
      sessionName: '',
      topicTitle: '',
      presentationAbstract: ''
    };

    // split the book section into an array
    let poster = section.split(/\n/);

    obj.sessionName = poster[0];
    obj.topicTitle = poster[1];
    obj.name = poster[2].split(', ');
    obj.affiliationName = poster[3];
    obj.presentationAbstract = poster.slice(4).join('\n');
	// Push separate object to JSON 
    dataArr.push(obj);
  });
  return dataArr;
}

pdf(dataBuffer).then(function(data) {
  
var xlsx = require('json-as-xlsx')

let columns = [
   {label: 'Name (incl. titles if any mentioned)', value: row => (row.name)},
   {label: 'Affiliation(s) Name(s)', value: row => (row.affiliationName)},
   {label: 'Persons Location', value: row => (row.personsLocation)},
   {label: 'Session Name', value: row => (row.sessionName)},
   {label: 'Topic Title', value: row => (row.topicTitle)},
   {label: 'Presentation Abstract', value: row => (row.presentationAbstract)},
 ]

 
 var settings = {
   sheetName: 'First sheet', // The name of the sheet
   fileName: 'Data_Entry', // The name of the spreadsheet
   extraLength: 5, // A bigger number means that columns should be wider
   writeOptions: {} // Style options from https://github.com/SheetJS/sheetjs#writing-options
 }
 
 var download = true // If true will download the xlsx file, otherwise will return a buffer
 
 xlsx(columns, splitText(data.text), settings, download) // Will download the excel file
  
});
