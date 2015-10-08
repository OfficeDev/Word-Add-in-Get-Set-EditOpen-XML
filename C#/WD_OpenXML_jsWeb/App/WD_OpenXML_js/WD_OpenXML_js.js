/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />

// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

      // Use this to check whether the new API is supported in the new Word Javascript API..
      if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
        // Do something that is only available via the new APIs
        
        $('#getOOXMLData').click(function () { getOOXML_newAPI(); });
        $('#setOOXMLData').click(function () { setOOXML_newAPI(); });
        console.log('This code is using Word 2016 or greater.');
      }
      else {
        // Just letting you know that this code will not work with your version of Word.
        console.log('This code is using Word 2013.');
        // Wire up the click events of the two buttons in the WD_OpenXML_js.html page.
        $('#getOOXMLData').click(function () { getOOXML(); });
        $('#setOOXMLData').click(function () { setOOXML(); });

      }



    });
};

// Variable to hold any Office Open XML
var currentOOXML = "";

function getOOXML_newAPI() {
  // Get a reference to the Div where we will write the status of our operation
  var report = document.getElementById("status");
  var textArea = document.getElementById("dataOOXML");
  // Remove all nodes from the status Div so we have a clean space to write to
  while (report.hasChildNodes()) {
    report.removeChild(report.lastChild);
  }

  // Run a batch operation against the Word object model.
  Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        currentOOXML = bodyOOXML.value;

        while (textArea.hasChildNodes()) {
          textArea.removeChild(textArea.lastChild);
          report.innerText = "";
        };
        setTimeout(function () {
          textArea.appendChild(document.createTextNode(currentOOXML));
          report.innerText = "The getOOXML function succeeded!";
        }, 400);

        // Clear the success message after a 2 second delay
        setTimeout(function () {
          report.innerText = "";
        }, 2000);

    });
  })
  .catch(function (error) {
      
      // Clear the OOXML, show the error info
          currentOOXML = "";
        report.innerText = error.message;    
  
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function setOOXML_newAPI() {
  // Get a reference to the Div where we will write the outcome of our operation
  var report = document.getElementById("status");

  //Sets the currentOOXML variable to the current contents of the task pane text area
  currentOOXML = document.getElementById("dataOOXML").textContent

  // Remove all nodes from the status Div so we have a clean space to write to
  while (report.hasChildNodes()) {
    report.removeChild(report.lastChild);
  }

  // Check whether we have OOXML in the variable
  if (currentOOXML != "") {

    // Call the setSelectedDataAsync, with parameters of:
    // 1. The Data to insert.
    // 2. The coercion type for that data.
    // 3. A callback function that lets us know if it succeeded.

    // Run a batch operation against the Word object model.
    Word.run(function (context) {

      // Create a proxy object for the document body.
      var body = context.document.body;

      // Queue a commmand to insert OOXML in to the beginning of the body.
      body.insertOoxml(currentOOXML, Word.InsertLocation.start);

      // Synchronize the document state by executing the queued commands, 
      // and return a promise to indicate task completion.
      return context.sync().then(function () {

        // Tell the user we succeeded and then clear the message after a 2 second delay
          report.innerText = "The setOOXML function succeeded!";
          setTimeout(function () {
            report.innerText = "";
          }, 2000);
      });
    })
    .catch(function (error) {

    // Clear the text area just so we don't give you the impression that there's
      // valid OOXML waiting to be inserted... 
      while (textArea.hasChildNodes()) {
        textArea.removeChild(textArea.lastChild);
      }
        // Let the user see the error.
        report.innerText = error.message;
        
        console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
      }
    });

  }
}

function getOOXML() {
    // Get a reference to the Div where we will write the status of our operation
    var report = document.getElementById("status");
    var textArea = document.getElementById("dataOOXML");
    // Remove all nodes from the status Div so we have a clean space to write to
    while (report.hasChildNodes()) {
        report.removeChild(report.lastChild);
    }

    // Now we can begin the process.
    // First we call the getSelectedDataAsync method. The included parameter is the coercion
    // type (in our case ooxml).
    // Note that the optional parameters valueFormat and filterType are not relevant to this
    // method when used in Word, so they are excluded here.
    // When the method returns, the function that is provided as the second parameter will run.
    Office.context.document.getSelectedDataAsync("ooxml",
        function (result) {
            // Get a reference to our textArea element,
            // which is located at the end of the Div with the ID 'Content' in the WD_OpenXML_js.html page.
            
            if (result.status == "succeeded") {

                // If the getSelectedDataAsync call succeeded, then
                // result.value will return a valid chunk of OOXML, which we'll
                // hold in the currentOOXML variable.

                currentOOXML = result.value;

                // Now we populate the text area in the task pane with the retrieved OOXML
                // so that you can copy it out for editing.
                //The first step below clears the text area and then we use a brief timeout to leave
                // the text area blank momentarily and make it clear that the OOXML is being refreshed
                // with the markup for the new selection.
                //Then we report to the user that we were successful

                while (textArea.hasChildNodes()) {
                    textArea.removeChild(textArea.lastChild);
                    report.innerText = "";
                };
                setTimeout(function () {
                    textArea.appendChild(document.createTextNode(currentOOXML));
                    report.innerText = "The getOOXML function succeeded!";
                }, 400);

                // Clear the success message after a 2 second delay
                setTimeout(function () {
                    report.innerText = "";
                }, 2000);
            }
            else {
                // This runs if the getSelectedDataAsync method does not return a success flag
                currentOOXML = "";
                report.innerText = result.error.message;
            }
        });
}

function setOOXML() {
    // Get a reference to the Div where we will write the outcome of our operation
    var report = document.getElementById("status");
    
    //Sets the currentOOXML variable to the current contents of the task pane text area
    currentOOXML = document.getElementById("dataOOXML").textContent

    // Remove all nodes from the status Div so we have a clean space to write to
    while (report.hasChildNodes()) {
        report.removeChild(report.lastChild);
    }

    // Check whether we have OOXML in the variable
    if (currentOOXML != "") {

        // Call the setSelectedDataAsync, with parameters of:
        // 1. The Data to insert.
        // 2. The coercion type for that data.
        // 3. A callback function that lets us know if it succeeded.

        
        Office.context.document.setSelectedDataAsync(
            currentOOXML, { coercionType: "ooxml" },
            function (result) {
                // Tell the user we succeeded and then clear the message after a 2 second delay
                if (result.status == "succeeded") {
                    report.innerText = "The setOOXML function succeeded!";
                    setTimeout(function () {
                        report.innerText = "";
                    }, 2000);
                }
                else {
                    // This runs if the getSliceAsync method does not return a success flag
                    report.innerText = result.error.message;

                    // Clear the text area just so we don't give you the impression that there's
                    // valid OOXML waiting to be inserted... 
                    while (textArea.hasChildNodes()) {
                        textArea.removeChild(textArea.lastChild);
                    }
                }
            });
    }
    else {

        // If currentOOXML == "" then we should not even try to insert it, because
        // that is gauranteed to cause an exception, needlessly.
        report.innerText = "There is currently no OOXML to insert!"
            + " Please select some of your document and click [Get OOXML] first!";
    }
}
// *********************************************************
//
// Word-Add-in-Get-Set-EditOpen-XML, https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
