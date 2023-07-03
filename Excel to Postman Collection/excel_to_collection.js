// Modules for reading .xlsx file
const reader = require('xlsx')
const readline = require("readline-sync");
const fs = require('node:fs');
// Module for creating postman collection
const { Collection, Item } = require('postman-collection');


function readXLSXFile() {
   //requires
   // none
   //guarantees
   // if(excel file exists)
   //    File is read.
   // else
   //    Display error

   console.log("Please enter the name of the xlsx file (without .xlsx extension):");
   var fileName = readline.question();
   fileName = fileName + ".xlsx";

   // Reading .xlsx file
   let file;
   try {
      file = reader.readFile(fileName)
   } catch (err) {
      if (err.code === 'ENOENT') {
         console.log(`${fileName} file is not found in the current directory!`);
         process.exit();
      } else {
         throw err;
      }
   }

   return file;

}

function createPostmanCollection(sheetName, sheetData) {
   //requires
   // Valid .xlsx sheet data and name.
   //guarantees
   // Collection file created.

   const postmanCollection = new Collection({
      info: {
         // Name of the collection
         name: sheetName
      },
      // Requests in this collection
      item: [],
   });
   sheetData.forEach((sheetElements) => {

      // Name of the request
      var requestName = sheetElements["Request Name"];

      // API endpoint
      var apiEndpoint = sheetElements["URL"];

      // Request method
      var method = sheetElements["Method"];

      // Request Authorization
      var requestAuthType = sheetElements["Authorization Type"]
      var requestAuth
      if (requestAuthType != null) {
         if (requestAuthType == "basic") {
            if (sheetElements["Username"] == null || sheetElements["Password"] == null) {
               console.log("Please provide username and password");
               process.exit();
            }
            requestAuth =
            {
               "type": "basic",
               "basic": [
                  { "key": "username", "value": sheetElements["Username"] },
                  { "key": "password", "value": sheetElements["Password"] }
               ]
            };
         } else if (requestAuthType == "bearer") {
            requestAuth =
            {
               "type": "bearer",
            };
         }
      }

      // Request Header
      var requestHeader
      if (sheetElements["Header"] != null)
         requestHeader = JSON.parse(sheetElements["Header"].replaceAll("'", '"'));

      // Request body
      var requestBody = sheetElements["Body"];


      const postmanRequest = new Item({
         name: `${requestName}`,
         request: {
            url: apiEndpoint,
            method: method,
            auth: requestAuth,
            header: requestHeader,
            body: requestBody
         }
      });

      // Add the reqest to our empty collection
      postmanCollection.items.add(postmanRequest);

   })

   // Convert the collection to JSON 
   // so that it can be exported to a file
   const collectionJSON = postmanCollection.toJSON();
   return collectionJSON;

}

function createPostmanCollectionFile(collectionJSON, sheetName) {
   //requires
   // Valid JSON Collection and sheet name.
   //guarantees
   // if(no error)
   //    Collection file created.
   // else
   //    Display error

   // Create a collection.json file. It can be imported to postman
   collectionFileName = sheetName + " collection.json";
   fs.writeFile(collectionFileName, JSON.stringify(collectionJSON), (err) => {
      if (err) { console.log(err); }
      console.log(sheetName + ' Collection has been created');
   });

}

//Main code starts here

//Reading .xlsx file
var file = readXLSXFile();

//Reading sheets
const sheets = file.SheetNames

//Parsing each sheet
for (let i = 0; i < sheets.length; i++) {
   const sheetData = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
   var collectionJSON = createPostmanCollection(file.SheetNames[i], sheetData);
   createPostmanCollectionFile(collectionJSON, file.SheetNames[i]);
}