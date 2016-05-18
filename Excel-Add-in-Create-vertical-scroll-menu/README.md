# Excel-Add-in-Create-vertical-scroll-men
This code sample demonstrates a task pane add-in that is displayed in Excel 2013 when the spreadsheet is opened.

The task pane contains a stacked tabbed menu along its right side. Moving the mouse over the menu causes the menu item under the mouse to move to the left. Moving the mouse off of the item causes it to move back to the right. Moving the mouse downwards causes the menu to vertically scroll up to reveal more menu options.

Figure 1 shows the task pane with the tabbed menu.

![Figure 1. Task pane with the tabbed menu](/description/image.jpg)


The sample demonstrates how to perform the following tasks:

* Attach event handlers to HTML elements.
* Use custom JQuery functions to animate HTML elements.
* Dynamically add style settings to HTML elements to change the display of the dashboard.

**Prerequisites**

This sample requires:

* Visual Studio 2012.
* Office Developer Tools for Visual Studio 2012.
* Excel 2013.

**Key components of the sample**

The sample app contains the following components:

* The VerticallyScrollingMenu project, which contains the VerticallyScrollingMenu.xml manifest file. The XML manifest file of an add-in for Office enables you to declaratively describe how the add-in should be activated when you install and use it with Office documents and applications.
* The VerticallyScrollingMenuWeb project, which contains multiple template files. However, the three files that have been developed as part of this sample solution include:

  * VerticallyScrollingMenu.html (in the Pages folder). This file contains the HTML user interface that is displayed in the task pane when the add-in is started. The markup consists of a <div> element containing a paragraph element which contains some random sample text. It also contains another <div> element that has the ID of  sidebar which contains an unordered list with an ID of  menu. The list contains a series of items consisting of an anchor element and a span element with some text. The list items are the tabbed menu options.
  * App.css (in the Styles folder). This cascading style sheet (CSS) contains the code that specifies the look of the sample text and the elements that make up the tabbed menu. Particularly notice the overflow:hidden; setting that causes any content that is greater than the top attribute setting of the sidebar div to be hidden.

```CSS
    #sidebar {
    height:400px;
    overflow:hidden;
    position:relative;
    background-color:#eee;
   }
   ```
   
   * VerticallyScrollingMenu.js (in the Scripts folder). This script file contains code that runs when the task pane add-in is loaded. Specifically, the script consists of commands from the JavaScript JQuery libraries named jquery-ui.js and jquery-1.9.1.min.js. This startup script first sets variables with the attributes for the tabbed menu when you move the mouse over each item.


```JavaScript 

      var colorOver = '#31b8da';
      var colorOut = '#1f1f1f';
      //Padding, mouseover
      var padLeft = '20px';
      var padRight = '20px'
      //Default Padding
      var defpadLeft = $('#menu li a').css('paddingLeft');
      var defpadRight = $('#menu li a').css('paddingRight');
   ```
   
The next code animates the tabbed menu when the mouse is moved over and off of the menu items. This animation causes the tabs to move to the left when the mouse moves over the item (the mouseover event) and back to the right when the mouse is moved off of the item (the mouseout event). The variables that you set previously are dynamically added to the list item's attributes to cause this effect.

```JavaScript 

$('#menu li').click(function () {
    //Make the LI clickable
    window.location = $(this).find('a').attr('href');

    }).mouseover(function () {

    //Mouse over the LI and look for an element 
    //for transition
    $(this).find('a')
    .animate({ paddingLeft: padLeft, paddingRight: padRight }, { queue: false, duration: 100 })
    .animate({ backgroundColor: colorOver }, { queue: false, duration: 200 });

    }).mouseout(function () {

    //Mouse out from the LI and look for an element 
    //and discard the mouse over transition
    $(this).find('a')
    .animate({ paddingLeft: defpadLeft, paddingRight: defpadRight }, { queue: false, duration: 100 })
    .animate({ backgroundColor: colorOut }, { queue: false, duration: 200 });
    });
 

   ```

The following code animates the menu so that it vertically scrolls up above the outer top limits of the sidebar div element. This makes it appear that the menu disappears as it moves upward.

```JavaScript 
$('#sidebar').mousemove(function (e) {

    //Sidebar Offset, Top value
    var s_top = parseInt($('#sidebar').offset().top);

    //Sidebar Offset, Bottom value
    var s_bottom = parseInt($('#sidebar').height() + s_top);

    //Roughly calculate the height of the menu by 
    //multiplying the height of a single LI with 
    //the total number of LIs
    var mheight = parseInt($('#menu li').height() * $('#menu li').length);

    //Calculate the top value
    var top_value = Math.round(((s_top - e.pageY) / 100) * mheight / 2)

    //Animate the #menu by changing the top value
    $('#menu').animate({ top: top_value }, { queue: false, duration: 500 });
    });
   ```
   

This code sample also requires the use of a custom JQuery library named jquery-ui.js that contains the functions that enable the menu animations. All other files are automatically provided by the Visual Studio project template for add-ins for Office, and they have not been modified in the development of this sample app.

**Configure the sample**

To configure the sample, open the VerticallyScrollingMenu.sln file with Visual Studio 2012. No other configuration is necessary.

**Build the sample**

To build the sample, choose Ctrl+Shift+B, or on the Build menu, select Build Solution.

**Run and test the sample**

To run the sample, choose the F5 key. After the task pane is displayed in Excel 2013, notice the stacked tabbed menu on the right side of the task pane. Moving the mouse over a particular tab menu causes the tab to move to the left. Moving the mouse off of the tabbed item moves the item back to the right. Moving the mouse down causes the menu to move upwards and moving the mouse upwards causes the menu to move down.

**Troubleshooting**

If the app fails to install, ensure that the XML in your AnimatedDashboard.xml manifest file parses correctly. Also look for any errors in the JavaScript code that could keep the tabbed menu from being displayed. For example, you may have forgotten to end a statement with a semicolon, or you may have misspelled a method name or keyword. If the components in the task pane do not look as you think they should, check the CSS styles to ensure that you didn't forget a colon between the style and its value, or leave off a semicolon at the end of a style statement.

**Change log**

* First release: April 29, 2013.
* GitHub release: August 13, 2015.

**Related content**

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Build apps for Office](http://msdn.microsoft.com/en-us/library/jj220060.aspx)
* [HTML Tutorial](http://www.w3schools.com/html/)
* [What is jQuery?](http://jquery.com/)
* [CSS Introduction](http://www.w3schools.com/css/css_intro.asp)

   
