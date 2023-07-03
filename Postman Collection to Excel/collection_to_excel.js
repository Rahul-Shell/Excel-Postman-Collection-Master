// Module for reading user input
const readline = require("readline-sync");
// Module for running collection in node js
const newman = require('newman');
// Module for creating .xlsx file
const XLSX = require('xlsx');

//-----------------WARNING------------------
//This is for endpoints that has incomplete certificate chain.
//Using this will make your requests insecure
//Please read https://stackoverflow.com/questions/31673587/error-unable-to-verify-the-first-certificate-in-nodejs
process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = 0;

function createXLSXFile(requestResponseData, collectionName) {
  //requires
  // Valid requestResponseData and collection name.
  //guarantees
  //  .xlsx file created.

  const workSheet = XLSX.utils.json_to_sheet(requestResponseData);
  const workBook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workBook, workSheet);

  excelName = collectionName + ".xlsx";
  XLSX.writeFile(workBook, excelName);
  console.log(`${excelName} file created`);

}

function generateRequestResponseData(fileName) {
  //requires
  // Valid fileName.
  //guarantees
  // if(no error)
  //  Creates requestResponseData.
  // else
  //    Display error

  newman.run({
    collection: require(`./${fileName}`),
  })
    .on('beforeDone', (error, data) => {
      if (error) {
        console.log(error);
        return;
      }

      requestResponseData = [];
      data.summary.run.executions.forEach(data => {
        const columns = {
          'Request Name': '',
          'Method': '',
          'URL': '',
          'Request Header': '',
          'Request Body': '',
          'Status': '',
          'Code': '',
          'Response Time': '',
          'Response Size': '',
          'Response Body': ''
        }

        //Request Data
        requestItem = data.item
        columns['Request Name'] = requestItem.name;
        columns['Method'] = requestItem.request.method;
        columns['URL'] = requestItem.request.url.toString();
        columns['Request Header'] = data.request.headers.members.toString();
        if (requestItem.request.body != null) {
          mode = requestItem.request.body.mode;
          columns['Request Body'] = requestItem.request.body[mode].toString();
        }

        //Response Data
        if(data.response!=null) {
          columns['Status'] = data.response.status;
          columns['Code'] = data.response.code;
          columns['Response Time'] = data.response.responseTime;
          columns['Response Size'] = data.response.responseSize;
          columns['Response Body'] = data.response.stream.toString();
        }

        requestResponseData.push(columns);
      });
      createXLSXFile(requestResponseData, data.summary.collection.name);
    });

}

//Main code starts here
console.log("Please enter the name of the collection file: ");
var fileName = readline.question();
fileName = fileName + ".json";

generateRequestResponseData(fileName);