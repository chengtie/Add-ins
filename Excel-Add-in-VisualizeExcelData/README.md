# Excel-Add-in-VisualizeExcelData
Sample shows how to create data visualizations in an Excel content add-in from table data in a spreadsheet

The Office sample includes a task pane add-in and a content add-in. It also includes an Excel workbook, PopulationVisualization.xlsx, that contains sample data.

The PopulationVisualization.xlsx file is set as the StartAction property of the add-in for Office. The workbook has a named cell called VisualStyle (cell B2), and a named table called PopulationTable (cells A6:B16). The following screen shots shows how the document and the add-in will appear after you start the solution. Figure 1 shows the workbook opened with the content add-in displayed.

![Figure 1](/description/CG_XLDataVis_fig01.gif)

Figure 2 shows the task pane add-in UI.

![Figure 2](/description/CG_XLDataVis_fig02.gif)


**The sample shows:**

* How to use JavaScript to add bindings to the named cell and the named table in the workbook.
* How to verify that those bindings are in place before attempting to retrieve the values from them or to set their values in code.
* How to set values in bindings from code, specifically in this case from the task pane add-in.
* How to retrieve values from the named cell and named table, and to use that data to create different types of visualizations for the data.
* How to respond to change events for the data bindings.


**Prerequisites**

This sample requires:

* Visual Studio 2012 (RTM).
* Office 2013 tools for Visual Studio 2012 (RTM).
* Excel 2013 (RTM).

**Key components of the sample**

The sample app contains:

* The ExcelDataVisualization project, which contains:
* The ExcelDataVisualization.xml manifest file.
* The ExcelDataVisualizationTaskPane.xml manifest file.
* The PopulationVisualization.xlsx document, which is prepopulated with a named cell and a named table.
* The ExcelDataVisualizationWeb project, which contains multiple template files. However, the four files that have been developed as part of this sample solution include:
* PopulationVisualizationTaskPane.html (in the Pages folder). This contains the HTML user interface that is displayed in the task pane. It consists of a <div> with an ID of validationReport, and two buttons.
* PopulationVisualizationTaskPane.html (in the Pages folder). This contains the HTML user interface that is displayed in the task pane. It consists of a <div> with an ID of validationReport, and two buttons.
* PopulationVisualizationContent.html (in the Pages folder). This contains the HTML user interface that is displayed in the task pane. It consists of a <div> with an IDof chart.
* PopulationVisualizationTaskPane.js (in the Scripts folder). This script file contains code that runs when the content add-in is loaded. This startup script attempts to add bindings to the named cell in the workbook. The success or failure of this operation is reported back to the PopulationVisualizationTaskPane.html page. The script file also includes the Click event handlers for the two buttons in PopulationVisualizationTaskPane.html. One of these buttons sets the value in the binding for the named cell to 'Stacked', and the other buttons sets the value in the binding for the named cell to 'Tiled'. These changes will raise binding events that are handled in PopulationVisualizationContent.js (see below).
* PopulationVisualizationContent.js (in the Scripts folder). This script file contains code that runs when the content add-in is loaded. This startup script attempts to add bindings to the named cell and named table in the workbook in the document. If the named cell has a value of either "Stacked" or "Tiled", the script builds an appropriate chart based on the data in the named table. Otherwise, it writes a simple message to the chart area, informing the user to choose a style for the visualization. The script file also includes two handlers, one of which responds to data change events in the binding to the named cell, and the other which responds to data change events in the binding to the named table. When either of these events is raised and handled, the current data visualization is destroyed and then recreated.

All other files are automatically provided by the Visual Studio project template for add-ins for Office, and they have not been modified in the development of this sample app.

**Configure the sample**

To configure the sample, open the ExcelDataVisualization.sln file with Visual Studio 2012. No other configuration is necessary.

**Build the sample**

To build the sample, choose the Ctrl+Shift+B keys.

**Run and test the sample**

To run the sample, choose the F5 key.

The following images show examples of the workbook at various stages of the process.

Figure 3 shows that the code has successfully created a binding to the named cell. The status of this initial binding has been reported in the task pane.

Figure 3. The task pane UI showing that the binding was successful.

![Figure 3](/description/CG_XLDataVis_fig03.gif)

The View stacked populations button calls code that takes the data in the data table and uses it to build the stacked visualization. In this case, squares representing each country's population have been overlaid on each other. When the user pauses the mouse pointer on a specific shade of green (representing a country), the tooltip provides the name of the country and the population value. Also note that the task pane includes a message that the visual style has been set successfully.

Figure 4 shows how the content pane appears when the View stacked populations button has been chosen.

Figure 4. The stacked populations view.

![Figure 4](/description/CG_XLDataVis_fig04.gif)

Alternatively, you could choose the View tiled populations. When you view the population in a tiled visualization, the task pane UI looks the same as when you choose the stacked visualization, but the content add-in changes to show the various populations in tiles. In this case, squares representing each country's population have been displayed in a jigsaw-like manner. When the user pauses the mouse pointer on a specific shade of green (representing a country), the tooltip provides the name of the country and the population value. Also note that the task pane includes a message that the visual style has been set successfully.

Figure 5 shows how the content pane appears when the View tiled populations button has been chosen.

Figure 5. The tiled populations view.

![Figure 5](/description/CG_XLDataVis_fig05.gif)

<a name="troubleshooting"></a>
**Troubleshooting**

If the add-in starts with a blank document instead of the one shown in Figure 1, ensure that the StartAction property of the ExcelDataVisualization project is set to PopulationVisualization.xlsx and not just to Excel.

**Change log**


* First release: March 15, 2013.
* Released on GitHub: August 13, 2015.

Related content


* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Build an add-in for Office](http://msdn.microsoft.com/en-us/library/jj220060.aspx)
* [Binding to regions in a document or spreadsheet](http://msdn.microsoft.com/en-us/library/fp123511.aspx)

