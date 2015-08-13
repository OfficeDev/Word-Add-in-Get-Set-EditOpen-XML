/// <reference path="../App.js" />

// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        // Wire up the click events of the two buttons in the WD_OpenXML_js.html page.
        $('#getOOXMLData').click(function () { getOOXML(); });
        $('#setOOXMLData').click(function () { setOOXML(); });
    });
};

// Variable to hold any Office Open XML
var currentOOXML = "";

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

