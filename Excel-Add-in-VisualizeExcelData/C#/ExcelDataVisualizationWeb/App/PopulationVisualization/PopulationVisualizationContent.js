/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />

(function () {
    "use strict";

    // This function is run when the app is ready to start interacting with the host application
    // It ensures the DOM is ready before binding to data in the worksheet
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // First main thing to do is create a binding to the named cell called VisualStyle.
            // So we'll start by getting a reference to the DIV in the PopulationVisualizationContent.html page, 
            // and then clearing its contents because we will write a message there if the binding fails
            var chart = document.getElementById("chart");
            var chartText;
            // Remove all nodes from the chart <DIV> so we have a clean space to write to
            while (chart.hasChildNodes()) {
                chart.removeChild(chart.lastChild);
            }

            // Bind to a named cell called VisualStyle. NOTE: This is cell A2 in the sample spreadsheet
            Office.context.document.bindings.addFromNamedItemAsync("VisualStyle", Office.BindingType.Text, { id: 'chartStyle' }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    // Write a message to the chart div if the binding has failed
                    chartText = document.createTextNode("Could not bind to VisualStyle named cell. Error Details are: " + asyncResult.error.message);
                    chart.appendChild(chartText);
                }
                else {
                    // If adding the binding has not failed, attempt to get it so we can then read it's value
                    Office.select("bindings#chartStyle").getDataAsync(function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            // Write a message to the chart div if retrieving the binding value has failed 
                            chartText = document.createTextNode("Error retrieving the Visual Style binding value: " + asyncResult.error.message);
                            chart.appendChild(chartText);
                        }
                        else {
                            // If getting the data succeeds, add an event handler so we can react to changes in the binding value
                            Office.select("bindings#chartStyle").addHandlerAsync(Office.EventType.BindingDataChanged, chartStyleChanged);
                            
                            // And then build the chart based on the binding value
                            buildChart(asyncResult.value);
                        }
                    });
                }
            });

        });
    };    
})();

// Event handler that we added in the $(document).ready function above
function chartStyleChanged(eventArgs) {
    // We are reacting to a change in the Visual Style binding, either as a result of the user clicking
    // a button in the task pane app, or by them tying directly in cell A2.
    // So effectively we are clearing the UI and then attempting to rebuild it again based on
    // the new setting.
    Office.select("bindings#chartStyle").getDataAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            // Write a message to the chart div if retrieving the binding value has failed 
            var chart = document.getElementById("chart");
            var chartText;
            // Remove all nodes from the chart <DIV> so we have a clean space to write to
            while (chart.hasChildNodes()) {
                chart.removeChild(chart.lastChild);
            }
            chartText = document.createTextNode("Error retrieving the Visual Style binding value: " + asyncResult.error.message);
            chart.appendChild(chartText);
        } else {
            // Build the chart based on the binding value
            buildChart(asyncResult.value);
        }
    });

}

// Event handler that we added in the $(document).ready function above
function chartDataChanged(eventArgs) {
    // Even though we are reacting to the actual chart data changing, we still need to get the Visual Style setting before we can re-build the
    // data visualization. NOTE: The buildChart function which is called by this handler automatically gets the latest data from the Population Table
    // so we don't need to do that here.
    Office.select("bindings#chartStyle").getDataAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            // Write a message to the chart div if retrieving the binding value has failed 
            var chart = document.getElementById("chart");
            var chartText;
            // Remove all nodes from the chart <DIV> so we have a clean space to write to
            while (chart.hasChildNodes()) {
                chart.removeChild(chart.lastChild);
            }
            chartText = document.createTextNode("Error retrieving the Visual Style binding value: " + asyncResult.error.message);
            chart.appendChild(chartText);
        } else {
            // Build the chart based on the binding value
            buildChart(asyncResult.value);
        }
    });

}

function buildChart(visualSetting) {
    //Remove all nodes from the chart <DIV> so we have a clean space to write to
    var chart = document.getElementById("chart");
    var chartText;
    while (chart.hasChildNodes()) {
        chart.removeChild(chart.lastChild);
    }

    // If the setting is neither Staked nor Tiled (e.g. when the document is first opened, or if a user types a different value in Cell A2)
    // then provide a hint to the user 
    if ((visualSetting != "Stacked") && (visualSetting != "Tiled")) {
        chartText = document.createTextNode("Please use the Population Visualization Settings task pane app to choose a visual style.");
        chart.appendChild(chartText);
        return;
    }

    // Otherwise, bind to the data in the named table called PopulationTable
    Office.context.document.bindings.addFromNamedItemAsync("PopulationTable", Office.BindingType.Table, { id: 'chartData' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            //Write a message to the chart div if the binding has failed
            chartText = document.createTextNode("Could not bind to PopulationTable data. Error Details are: " + asyncResult.error.message);
            chart.appendChild(chartText);
        }
        else {
            // If adding the binding has not failed, add an event handler so we can react to changes in the table values
            Office.select("bindings#chartData").addHandlerAsync(Office.EventType.BindingDataChanged, chartDataChanged);

            //Then attempt to get it so we can then read its data.
            // In this case, the data will be an array of country names and populations
            Office.select("bindings#chartData").getDataAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    //Write a message to the chart div if retrieving the binding value has failed 
                    chartText = document.createTextNode("Error retrieving the PopulationTable data values: " + asyncResult.error.message);
                    chart.appendChild(chartText);
                } else {
                    // If the PopulationTable has been read successfully, build a chart based on the Visual Style setting.
                    // First this here is to read the data from the TableData object into our own array, for easier manipulation.
                    var countryTableData = asyncResult.value;
                    var countryData = [];
                    for (country = 0; country < countryTableData.rows.length; country++) {
                        countryData.push({ "countryName": countryTableData.rows[country][0], "population": countryTableData.rows[country][1], });
                    }

                    // Take a copy of our array, simply so that we can sort it and then find the largest population.
                    // The largest population is used later so that we can scale the areas of all the other smaller countrys
                    // against the largest one
                    var tempData = countryData.slice(0);
                    tempData.sort(sort_by('population', true, parseInt));
                    var largestPopulation = tempData[0].population;

                    // If we are going to show the data stacked, then we will now sort the actual data so that the countries
                    // with the largest populations are added to the chart first. This is so they can contain the successively smaller
                    // countries. NOTE: The reason we don't want to sort for the Tiled UI is that this looks better with unordered tiles.
                    if (visualSetting == "Stacked") {
                        countryData.sort(sort_by('population', true, parseInt));
                    }

                    // Clear the chart area, and then start adding sqaures to represent each country's population
                    var visualArea = document.getElementById("chart");
                    while (visualArea.hasChildNodes()) {
                        visualArea.removeChild(visualArea.lastChild);
                    }
                    var countryCount = countryData.length;
                    for (var country = 0; country < countryCount; country++) {
                        var proportionalSize;
                        if (visualSetting == "Stacked") {
                            // We have already determined the largest population, so all countries can now be built such that
                            // their area in pixels is relative to that of the largest country.
                            // Of course, because we are building squares, the width and height need to be the square root of the relative
                            // proportion, so that the final square has an area that relates properly to the largest country's square.
                            // The factor of 80000 is simply so that squares are of a reasonable size
                            proportionalSize = Math.sqrt((countryData[country].population / largestPopulation) * 80000);
                        }
                        else {
                            // Same calculation as above, but the factor of 25000 is because we want smaller tiles than above,
                            //so that we more can fit in the visible area of the chart
                            proportionalSize = Math.sqrt((countryData[country].population / largestPopulation) * 25000);

                        }
                        // The following div is going to be a square that represents the country's population by its area size
                        var countryDiv = document.createElement("div");
                        if (visualSetting == "Stacked") {
                            // Set styles to achieve two main things:
                            // 1. Make the div into a square with sides equal to the proportionalSize variable that we determined above. 
                            //    This will mean that the area of the div represents the population (i.e. width*height)
                            // 2. Set the background color of the square to be a different shade of green than all the other countries.
                            //    Effectively, we start with a base color of RGB(000,100,000) and then assign a green component that
                            //    is on a point equally spaced between 100 and 255, depending on how many countries there are.
                            //    By the way, using the .toString(16) method converts the decimal number into a hex value, which is 
                            //    perfect for our needs.
                            countryDiv.setAttribute("style", "padding:2px;float:left;width:"
                                + proportionalSize + "px;height:"
                                + proportionalSize + "px;background-color:#00"
                                + (100 + (parseInt((155 / countryCount) * country))).toString(16) + "00");
                        }
                        else {
                            // Same as above, execpt we don't want the padding for when the squares are tiled
                            countryDiv.setAttribute("style", "float:left;width:" + proportionalSize + "px;height:" + proportionalSize + "px;background-color:#00" + (100 + (parseInt((155 / countryCount) * country))).toString(16) + "00");
                        }

                        // Set the tooltip of the div so that it shows both the country and the population number when the mouse hovers over the square
                        countryDiv.setAttribute("title", countryData[country].countryName + " - " + countryData[country].population);
                        visualArea.appendChild(countryDiv);
                        // On the first iteration, the visualArea variable is set to the actual chart div in the content app page.
                        // If we are going to tile, then that's OK --- we'll just keep adding new floating divs, and the app will arrange them for us
                        // depending on the size of the area.
                        // However, if we are going to stack the divs inside each other, then we simply need to set the visualArea variable
                        // to the current div. That way, when the next iteration occurs, the previous div will become the place where the new div gets added.
                        if (visualSetting == "Stacked") {
                            visualArea = countryDiv;
                        }
                    }
                }
            });

        }
    });
}

// Helper function for sorting the data in our array in reverse order
var sort_by = function (field, reverse, primer) {
    var key = function (country) { return primer ? primer(country[field]) : country[field] };
    return function (a, b) {
        var A = key(a), B = key(b);
        return ((A < B) ? -1 : (A > B) ? +1 : 0) * [-1, 1][+!reverse];
    }
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