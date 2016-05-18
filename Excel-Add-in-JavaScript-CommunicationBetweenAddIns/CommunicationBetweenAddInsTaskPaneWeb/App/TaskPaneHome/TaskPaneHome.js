/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

// Declare global variables for storing user data.
var percentage;
var downPayment;
var loanTerm;

// This function is run when the add-in is ready to start interacting with the host application.
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        // Remove data from localStorage on unload.
        window.onunload = function () {
            localStorage.removeItem("percentage");
            localStorage.removeItem("downpayment");
            localStorage.removeItem("loanterm");
        }
    });
};

// NOTE: If you are using live data to set the values communicated to the other add-in
// for Office, you can set an interval to get and store the data on an intermittent
// basis.
//var providerInterval;

//function startTimer() {
//    providerInterval = setInterval(setValue, 100);

//}

// Set the current values to the localStorage.
function setValue() {

    try {
        // Get the user-selected values from the interface.
        percentage = document.getElementById('percentagerate').value;
        downPayment = document.getElementById('downpayment').value;

        var loanTermList = document.getElementById('loanterm');
        var index = loanTermList.selectedIndex;
        var termSelected = loanTermList.options[index];
        loanTerm = termSelected.value;

        // Define a function to make sure that the percentage
        // is a number between 0 and 10.
        var percentValid = function (value) {
            return (Number(value) < 10) &
                   (Number(value) > 0) &
                   !isNaN(value);
        }

        // Define a function to make sure that the down payment 
        // is a number greater than $1,000.00.
        var downValid = function (value) {
            return (Number(value) > 1000) &
                    !isNaN(value);
        }

        // Validate the user-entered values.
        if (validate(percentage, percentValid) &
            validate(downPayment, downValid)) {

            // Values are valid; update the data source.
            localStorage.setItem('percentage', percentage);
            localStorage.setItem('downpayment', downPayment);
            localStorage.setItem('loanterm', loanTerm);
        }
        else {
            var dataError = { name: "Data error", message: "Please check your data to make sure that it is correct." };
            throw dataError;
        }
    }
    catch (err) {
        Toast.showToast(err.name, err.message);
    }
}

// Check the value to make sure that it is within the accepted range.
function validate(value, compareFunction) {

    if (compareFunction(value)) {
        return true;
    }
    else {

        Toast.showToast("Data entry", value + " is not a valid value. <br />Please enter a valid value.");
        return false;
    }

}

// *********************************************************
//
// Excel-Add-in-JavaScript-CommunicationBetweenAddIns, https://github.com/OfficeDev/Excel-Add-in-JavaScript-CommunicationBetweenAddIns/
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
// ***********************************************************