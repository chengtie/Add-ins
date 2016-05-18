# Excel-Add-in-Display-Animated-Dashboard
This code sample demonstrates a task pane add-in that is displayed in Excel 2013 when the add-in is first started. The task pane contains a partially hidden menu along its left side. Moving the mouse over the dashboard causes the menu to be fully displayed. Each menu item also includes a button that is used to either insert the text from a text box into the worksheet or retrieve text from the worksheet and insert it into the text box.

Figure 1 shows the task pane with the partially displayed dashboard.

![Figure 1. Initial state of the dashboard](/description/image.jpg)

Figure 2 shows the fully displayed dashboard.

![Figure 2. Fully displayed dashboard](/description/42e94672-4944-4185-89b2-7b947132b088image.jpg)

 
The sample demonstrates how to perform the following tasks:

* Attach event handlers to HTML elements.
* Use custom JQuery functions to animate HTML elements.
* Dynamically add style settings to HTML elements to change the display of the dashboard.
* Retrieve selected content from the worksheet.
* Insert content from a text box into the worksheet.

*Prerequisites*

This sample requires:

* Visual Studio 2012.
* Office Developer Tools for Visual Studio 2012.
* Excel 2013.

*Key components of the sample*

The sample app contains the following components:

* The AnimatedDashboard project, which contains the AnimatedDashboard.xml manifest file. The XML manifest file of an add-in for Office enables you to declaratively describe how the add-in should be activated when you install and use it with Office documents and applications.
* The AnimatedDashboardWeb project, which contains multiple template files. However, the three files that have been developed as part of this sample solution include:
* AnimatedDashboard.html (in the Pages folder). This file contains the HTML user interface that is displayed in the task pane when the add-in is started. The markup consists of a <div> element that contains a text box element that has an ID of selectedDataTxt. It also contains another <div> element that has the ID of dashboard that contains two buttons that have IDs of  setDataBtn and getDataBtn. The  setDataBtn button inserts text from the text box into the worksheet. The  getDataBtn button retrieves any selected text from the worksheet and inserts it into the text box.
* App.css (in the Styles folder). This cascading style sheet (CSS) contains the code that specifies the initial look of the dashboard and the elements each menu item contains as shown in the following code. Particularly notice the left: -92px setting that causes the dashboard to appear partially hidden on the left side of the task pane.

```CSS
#dashboard {
width: 70px;
background-color: rgb(110,138,195);
padding: 20px 20px 20px 20px;
position: absolute;
left: -92px;
z-index: 100;
}
``` 

The CSS also contains the style code that specifies the appearance of the two buttons.

```CSS
#setDataBtn
{
margin-right: 10px; 
padding: 0px; 
width: 90px;
}

#getDataBtn
{
padding: 0px; 
width: 90px;
}
``` 

Finally, the following code formats the text box.

```CSS
#selectedDataTxt
{
margin-top: 10px; 
width: 210px
}
``` 
 

* AnimatedDashboard.js (in the Scripts folder). This script file contains code that runs when the task pane add-in is loaded. Specifically, the script consists of commands from the JavaScript JQuery libraries named jquery.easing.1.3.js and jquery-1.9.1.min.js. This startup script first attaches code to the hover event of the <div> element that has the ID dashboard that contains the menu items. The hover event takes two arguments that define what happens when you move the mouse over the menu and then what happens when you move the mouse off of the menu.

```JavaScript 
$('#dashboard').hover(
``` 

When the mouse is moved over the dashboard, the CSS left attribute value is dynamically changed from a negative value (menu is partially hidden) to 0, which causes the dashboard to be displayed. Next, the code sets the duration for the animation to 500 milliseconds, which equates to half a second. Finally, the code sets a custom easing method from the jquery.easing.1.3.js library that causes the dashboard to appear slowly at first and then speed up.



```JavaScript 

function() {
$(this).stop().animate(
{
    left: '0',
},
500,
'easeInSine'
);
``` 

When the mouse is moved off of the dashboard, the following code is run. The  left attribute is reset to a negative value, which causes it to be partially hidden on the left side of the task pane. Next, the code sets the duration for the animation to 1500 milliseconds, which equates to one and a half seconds. Finally the code sets a custom easing, which causes the dashboard to retract to the left and then appear to bounce before settling in.

```JavaScript 

function() {
$(this).stop().animate(
{
    left: '-92px'
},
1500,
'easeOutBounce'
);
``` 

When the Set data button is clicked, the click event is activated to call the setData function, passing in the text in the text box.

```JavaScript 

$('#setDataBtn').click(function () { setData('#selectedDataTxt'); });
 ```

The setData function calls the setSelectedDataAsync method to insert the text from the active panel into the worksheet. The setSelectedDataAsync method asynchronously writes data to the current selection in the document.



```JavaScript 
function setData(elementId) {
    Office.context.document.setSelectedDataAsync($(elementId).val());
}
``` 

Similar to the Set data button, the  Get data button activates the click event to call the getData function, passing in the ID of the selectedDataTxt text box.

```JavaScript 

$('#getDataBtn').click(function () { getData('#selectedDataTxt'); });

``` 

The getData function reads the data from current selection of the document and displays it in a text box



```JavaScript 

function getData(elementId) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    function (result) {
        if (result.status === 'succeeded') {
            $(elementId).val(result.value);
        }
});
 
```


This code sample also requires the use of a custom JQuery library named jquery.easing.1.3.js that contains the functions that enable the dashboard animations. All other files are automatically provided by the Visual Studio project template for apps for Office, and they have not been modified in the development of this sample app.

*Configure the sample*

To configure the sample, open the AnimatedDashboard.sln file with Visual Studio 2012. No other configuration is necessary.

*Build the sample*

To build the sample, choose Ctrl+Shift+B, or on the Build menu, select Build Solution.

*Run and test the sample*

To run the sample, choose the F5 key. After the task pane is displayed in Excel 2013, notice that there is a text box that contains sample text. There is also a dashboard on the left side of the task pane. Moving the mouse over the dashboard causes it to move to the right, displaying the Set data and  Get data buttons. Click the Set data button. Notice that the text from the text box is inserted into the worksheet. Change the text in the worksheet and then click the Get data button. Notice that the updated text appears in the text box.

*Troubleshooting*

If the add-in fails to install, ensure that the XML in your AnimatedDashboard.xml manifest file parses correctly. Also look for any errors in the JavaScript code that could keep the dashboard from being displayed. For example, you may have forgotten to end a statement with a semicolon, or you may have misspelled a method name or keyword. If the components in the task pane do not look as you think they should, check the CSS styles to ensure that you didn't forget a colon between the style and its value, or leave off a semicolon at the end of a style statement.

*Change log*

* First release: April 29, 2013.
* GitHub releas: August 20, 2015.

*Related content*

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Build apps for Office](http://msdn.microsoft.com/library/jj220060.aspx)
* [HTML Tutorial](http://www.w3schools.com/html/)
* [What is jQuery?](http://jquery.com/)
* [CSS Introduction](http://www.w3schools.com/css/css_intro.asp)
* [Document.setSelectedDataAsync method](http://msdn.microsoft.com/library/office/apps/fp142145.aspx)
* [Document.getSelectedDataAsync method](http://msdn.microsoft.com/library/office/apps/fp142294.aspx)


