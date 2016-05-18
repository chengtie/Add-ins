# Excel-Add-in-Create-tab-Menu
This code sample demonstrates a task pane add-in that is displayed in Excel 2013 when the add-in is first started. The task pane contains three tabs that are presented horizontally, each with a tab panel that contains some random text. Each tab also includes a button that is used to insert the text just from that tab into the worksheet.


Figure 1 shows the task pane with the three-tab menu displayed.

![Figure 1. Tab menu](/description/image.jpg)

 
The sample demonstrates how to perform the following tasks:

* Attach event handlers to HTML elements.
* Use JavaScript to hide HTML elements in the task pane.
* Dynamically add style settings to HTML elements to display the tab menu at a particular location on the screen.
* Use descendent qualifiers to select HTML elements.
* Insert content into the worksheet.

*Prerequisites*

This sample requires:

* Visual Studio 2012.
* Office Developer Tools for Visual Studio 2012
* Excel 2013.

*Key components of the sample*

The sample app contains the following components:

* The TabMenus project, which contains the TabMenus.xml manifest file. The XML manifest file of an add-in for Office enables you to declaratively describe how the add-in should be activated when you install and use it with Office documents and applications.
* The TabMenusWeb project, which contains multiple template files. However, the three files that have been developed as part of this sample solution include:
* TabMenus.html (in the Pages folder). This file contains the HTML user interface that is displayed in the task pane when the add-in is started. The markup consists of a <ul> (unordered list) element with a class name of tabs, where each <li> (list item) element is a tab in the menu. It also contains three <div> elements that have the IDs of panel1, panel2, and  panel3, which are the individual panels, each of which contains random text. It also contains an <input> element of type button that inserts the text from a particular panel into the worksheet when the button is chosen.
* App.css (in the Styles folder). This cascading style sheet (CSS) contains the code that specifies the look of the tabs and the elements each tab contains, as shown in the following code. Particularly notice the display: block setting that causes the tabs to appear horizontally.

```CSS
.tabs {
margin: 0;
padding: 0;
zoom : 1;
}
.tabs li {
float: left;
list-style: none;
padding: 0;
margin: 0;
}
.tabs a {
display: block;
text-decoration: none;
padding: 3px 5px;
background-color:aqua;
margin-right: 10px;
border: 1px solid rgb(153,153,153);
margin-bottom: -1px;
}
``` 

The CSS also contains the style code that changes the appearance of the tab when it becomes the active tab.

```CSS
.tabs .active {
border-bottom: 1px solid white;
background-color: white;
color: rgb(51,72,115);
position: relative;
``` 

The following style code defines the default appearance of each panel.

```CSS

.panelContainer {
clear: both;
margin-bottom: 25px;
border: 1px solid rgb(153,153,153);
background-color: white;
padding: 10px;

``` 

Finally, the following code formats the Insert Data button.

```CSS

#setDataBtn {
    margin-right: 10px; 
    padding: 0px; 
    width: 100px;
 
```

* TabMenus.js (in the Scripts folder). This script file contains code that runs when the task pane add-in is loaded. Specifically, the script consists of commands from the JavaScript JQuery library. This startup script displays the first tab and panel.

```JavaScript 

$('.tabs li:first a').click();
``` 

When a tab is chosen, the script attaches a click event to the anchor tags in the tab menu, which is then executed.

```JavaScript 

$('.tabs a').click(function ()
``` 

The script then stores the calling object (represented by the this keyword) into a variable so that it can be used later.

```JavaScript 

var $this = $(this);
``` 

Next, the script hides the panels and then clears the active state from the tabs. Then, the calling tab is set as active and its URL is retrieved and stored in the panel variable. The URL will be the address on the panel that is associated with that tab. Finally, the panel is relatively slowly faded in until it is displayed. The number passed in to the method is the duration, in milliseconds, of the animation.

```JavaScript 

$('.panel').hide();
$('.tabs a.active').removeClass('active');

$this.addClass('active').blur();

panel = $this.attr('href'); 

$(panel).fadeIn(250);
```
 
When the Insert data button is chosen, the click event is activated to call the setData function, passing in the paragraph of text that is associated with the active tab.

```JavaScript 

$('#setDataBtn').click(function () { setData($(panel).children('p')); });
``` 

The setData function calls the setSelectedDataAsync method to insert the text from the active panel into the worksheet. The setSelectedDataAsync method asynchronously writes data to the current selection in the document.

```JavaScript 

function setData(elementId) {
Office.context.document.setSelectedDataAsync($(elementId).text());
}
``` 

All other files are automatically provided by the Visual Studio project template for add-ins for Office, and they have not been modified in the development of this sample app.

*Configure the sample*

To configure the sample, open the TabMenus.sln file with Visual Studio. No other configuration is necessary.

*Build the sample*

To build the sample, choose Ctrl+Shift+B, or on the Build menu, select Build Solution.

*Run and test the sample*

To run the sample, choose the F5 key. After the task pane is displayed in Excel 2013, notice that there are three tabbed panels that contain text. Choosing each tab displays a different panel. Choose the Insert data button. Notice the text that is inserted into the worksheet. Select another panel and then choose the Insert data button. Notice that the text that is inserted into the worksheet changes.

*Troubleshooting*

If the add-in fails to install, ensure that the XML in your TabMenus.xml manifest file parses correctly. Also look for any errors in the JavaScript code that could keep the tabs from being displayed. For example, you may have forgotten to end a statement with a semicolon, or you may have misspelled a method name or keyword. If the tabs and panels in the task pane do not look as you think they should, check the CSS styles to ensure that you didn't forget a colon between the style and its value, or leave off a semicolon at the end of a style statement.

*Change log*

* First release: April 29, 2013.
* GitHub release: August 20, 2015.

*Related content*

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Build apps for Office](http://msdn.microsoft.com/library/jj220060.aspx)
* [HTML Tutorial](http://www.w3schools.com/html/)
* [What is jQuery?](http://jquery.com/)
* [CSS Introduction](http://www.w3schools.com/css/css_intro.asp)
* [Document.setSelectedDataAsync method](http://msdn.microsoft.com/library/office/apps/fp142145.aspx)

