/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />

(function () {
    "use strict";

    Office.initialize = function (reason) {
        // This function is run when the app is ready to start interacting with the host application
        // It ensures the DOM is ready before adding click handlers to buttons binding to a named cell
        $(document).ready(function () {

            checkSampleDocLoaded();
            // Ready the app object so that you can use its notification method using the 'app.showNotification' method
            app.initialize();
            // Wire up the click events of the two buttons in the PopulationVisualizationTaskPane.html page.
            $('#setStacked').click(function () { stack(); });
            $('#setTiled').click(function () { tile(); });
        });
        // Bind to a named cell called VisualStyle. NOTE: This is cell A2 in the sample spreadsheet
        Office.context.document.bindings.addFromNamedItemAsync("VisualStyle", Office.BindingType.Text, { id: 'visualStyle' }, function (asyncResult) {
            // Provide status information about the binding. NOTE: In a real solution, you wouldn't bother showing this to the user
            // but it can help you troubleshoot later if the binding cannot be found.
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification('Error', 'Could not bind to VisualStyle named cell. Error Details are: ' + asyncResult.error.message);
            }
            else {
                app.showNotification('Status', 'Successfully bound to named cell VisualStyle!');
            }
        });
    };
})();

// This function is wired up as the click event handler in the initialize method above
function stack() {
    // Attempt to get the nameBinding that should have been added in the $(document).ready function
    Office.context.document.bindings.getByIdAsync('visualStyle', function (asyncResult) {
        // If retrieving the binding fails, tell the developer
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            app.showNotification('Error', 'Could not retrieve the Visual Style binding');
            return;
        }
        else {
            //If retrieving the binding succeeds, attempt to set the actual data to 'Stacked'            
            Office.select("bindings#visualStyle").setDataAsync("Stacked", function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    //If setting the data fails, tell the developer
                    app.showNotification('Error', 'Could not set the Visual Style binding: ' + asyncResult.error.message);
                } else {
                    // If setting the data succeeds, tell the developer.
                    app.showNotification('Status', 'Visual style successfully set to Stacked!');
                }
            });
        }
    });
}

// This function is wired up as the click event handler in the initialize method at the top of this page
function tile() {
    // Attempt to get the nameBinding that should have been added in the $(document).ready function
    Office.context.document.bindings.getByIdAsync('visualStyle', function (asyncResult) {
        // If retrieving the binding fails, tell the developer
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            app.showNotification('Error', 'Could not retrieve the Visual Style binding');
            return;
        }
        else {
            //If retrieving the binding succeeds, attempt to set the actual data to 'Tiled'            
            Office.select("bindings#visualStyle").setDataAsync("Tiled", function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    //If setting the data fails, tell the developer
                    app.showNotification('Error', 'Could not set the Visual Style binding: ' + asyncResult.error.message);
                } else {
                    // If setting the data succeeds, tell the developer.
                    app.showNotification('Status', 'Visual Style successfully set to Tiled!');
                }
            });
        }
    });
}
function checkSampleDocLoaded() {

    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;

        // Find the warning section on the page. We'll hide or show it
        // depending on whether the sample doc is loaded or not. 
        var elem = document.getElementById('warning');

        // Note: If you change the name of the sample doc in the InvoiceManager.csproj, don't
        // forget to update it here 
        if (fileUrl == "" || fileUrl.indexOf("PopulationVisualization.xlsx") == -1) {
            elem.style.display = 'block'; // show warning
        }
        else {
            elem.style.display = 'none'; // hide warning
        }
    });
}
// *********************************************************
//
// Excel-Add-in-VisualizeExcelData, https://github.com/OfficeDev/Excel-Add-in-VisualizeExcelData
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