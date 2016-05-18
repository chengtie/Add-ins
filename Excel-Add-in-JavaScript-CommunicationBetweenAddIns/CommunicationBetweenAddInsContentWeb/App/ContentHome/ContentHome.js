/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

// Declare global variables for the add-in.
var percentage;
var downPayment;
var loanTerm;
var salePrices;

// This function is run when the add-in is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        // Add a handler to detect the DocumentSelectionChanged event.
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            selectionChanged);

        // Set initial values for the percentage, downPayment, and loanTerm global variables.
        percentage = 0;
        downPayment = 0;
        loanTerm = 0;
        salePrices = [];
    });
};

// Detect changes in the spreadsheet.
function selectionChanged() {

    // Get the new selected values in the spreadsheet.
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // Do something with the errors.
            }
            else {
                // Get selected data.
                var dataValue = asyncResult.value;
                if (dataValue != "") {

                    // Reinitialize the salePrices array and 
                    // assign the results to the array.
                    salePrices = [];

                    salePrices = dataValue;
                    displayValues();
                }
            }
        });

}

// Check for changes in the data emitted by the data provider.
function connectToData() {
    var interval = setInterval(getData, 500);
}

// Work with the data saved in localStorage by the connected Office add-in.
function getData() {

    var footer = document.getElementById("footer");
    var output = document.getElementById("output");
    var connect = document.getElementById("connect");

    // Make sure that the data is available.
    if ((localStorage.getItem("percentage") != null) &
        (localStorage.getItem("downpayment") != null) &
        (localStorage.getItem("loanterm") != null)) {

        // Check to see if the data has changed.
        if ((Number(localStorage.getItem("percentage")) != percentage) |
            (Number(localStorage.getItem("downpayment")) != downPayment) |
            (Number(localStorage.getItem("loanterm")) != loanTerm)) {

            // Get the data from localStorage.
            percentage = localStorage.getItem("percentage");
            downPayment = localStorage.getItem("downpayment");
            loanTerm = localStorage.getItem("loanterm");

            // Display the data from localStorage in the content add-in.
            document.getElementById("percentage").innerText = percentage;
            document.getElementById("loanterm").innerText = loanTerm;
            document.getElementById("downpayment").innerText = downPayment;

            // Toggle the visibility of the footer and output divs.
            output.setAttribute("style", "display:none");
            footer.setAttribute("style", "display:block");
            connect.value = "Connected ...";
            connect.setAttribute("disabled", "disabled");

            // Generate new values to display in the table if 
            // any sale prices have been selected.
            if (salePrices.length > 0) {
                displayValues();
            }
        }
    }
    else {

        // Display an error message in the output div.
        footer.setAttribute("style", "display:none");
        output.setAttribute("style", "display:block");
        output.innerText = "Click the submit button to test the connection.";

    }
}

// Create new values in the table.
function displayValues() {

    // Clear all the rows in the table.
    clearTable();

    // Enable scrolling if there are more than five items in the
    // salePrices array.
    if (salePrices.length > 5) {
        document.getElementById("display").setAttribute("class", "scrolling");
    }
    else {
        document.getElementById("display").setAttribute("class", "static");
    }

    // Iterate over the values in the salePrice array and 
    // create a new row in the table for each value.
    for (var i = 0; i < salePrices.length; i++) {

        // Get the sale price from the salePrices array and format it.
        var newPrice = salePrices[i];
        var loanAmount = newPrice - downPayment;
        loanAmount = loanAmount.toFixed(2);

        // Calculate the monthly payment from the data.
        var monthlyPayment = calculatePayment(loanAmount, loanTerm, percentage);

        // Create a new row for the data to display the loan amount and
        // monthly mortgage payment.
        if (monthlyPayment > 0) {
            createRow(loanAmount, monthlyPayment);
        }
    }
}

// Removes all existing rows from the table of displayed values.
function clearTable() {
    var tableBody = document.getElementById("tablebody");
    while (tableBody.hasChildNodes()) {
        tableBody.removeChild(tableBody.lastChild);
    }
}

// Create a new row in the table of displayed values.
function createRow(loanAmount, monthlyPayment) {
    var tableBody = document.getElementById("tablebody");

    // Create a new row element.
    var newRow = document.createElement("tr");

    // Create a cell for the loan Amount.
    var loanCell = document.createElement("td");
    var loanText = document.createTextNode("$" + loanAmount.toString());
    loanCell.appendChild(loanText);

    // Create a cell for the monthly payment.
    var paymentCell = document.createElement("td");
    var paymentText = document.createTextNode("$" + monthlyPayment.toString());
    paymentCell.appendChild(paymentText);

    //Add the new cells to the row and the new row to the table.
    newRow.appendChild(loanCell);
    newRow.appendChild(paymentCell);
    tableBody.appendChild(newRow);
}
// *********************************************************
//
// Excel-Add-in-JavaScript-CommunicateBetweenAddIns,https://github.com/OfficeDev/Excel-Add-in-JavaScript-CommunicationBetweenAddIns/
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