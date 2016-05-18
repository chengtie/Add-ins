# Excel-Add-in-Create-News-Ticker-Display
This sample app for Office shows how to use cascading style sheets and JavaScript in an Excel 2013 task pane add-in to create a news ticker type display.

This code sample demonstrates a task pane add-in that is displayed in Excel 2013 when the add-in is first started. The task pane displays a number of lines of text (with anchor elements) that scroll upward. As the top line of text disappears, another line of text appears at the bottom of the task pane. This type of action is suited to displaying news headlines from RSS feeds or other information that scrolls when size prohibits more details.


Figure 1 shows the task pane with the lines of news headlines.

![Figure 1. List of news headlines](/description/image.jpg)


The sample demonstrates how to perform the following tasks:

* Attach event handlers to HTML elements.
* Use custom JQuery functions to animate HTML elements.
* Dynamically add style settings to HTML elements to change the display of the list.
* Chaining JQuery functions together to make the code more efficient.

*Prerequisites*


This sample requires:

* Visual Studio 2012, 2013, or 2015.
* Office Developer Tools for Visual Studio.
* Excel 2013.

*Key components of the sample*

The sample app contains the following components:

* The News_Ticker project, which contains the News_Ticker.xml manifest file. The XML manifest file of an add-in for Office enables you to declaratively describe how the add-in should be activated when you install and use it with Office documents and applications.
* The News_TickerWeb project, which contains multiple template files. However, the three files that have been developed as part of this sample solution include:
* News_Ticker.html (in the Pages folder). This file contains the HTML user interface that is displayed in the task pane when the app is started. The markup consists of an unordered list element with the ID of listticker that contains a list of anchor elements enclosed in paragraph elements. There is also an animated image that precedes each anchor element.

```HTML 

<ul id="listticker">
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 1</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 2</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 3</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 4</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 5</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 6</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 7</a></p></li>
    <li><p><img src="../images/news.gif" /><a href="#">News Headline 8</a></p></li>
</ul>
``` 

* App.css (in the Styles folder). This cascading style sheet (CSS) contains the code that specifies the look of each line of the unordered list and the image.

```CSS
#listticker {
font-weight:900;
font-size: 20px;
}

#listticker li a {
color: blue;
}

img {
display: inline;
width: 25px;
height: 22px;
padding: 0 10px 0 0; 
}#dashboard {
width: 70px;
background-color: rgb(110,138,195);
padding: 20px 20px 20px 20px;
position: absolute;
left: -92px;
z-index: 100;
}
``` 


* News_Ticker.js (in the Scripts folder). This script file contains code that runs when the task pane add-in is loaded. Specifically, the script consists of commands from the JavaScript JQuery library named jquery-1.10.1.min.js. This code first sets variables that determine various intervals such as that used by the first item as it fades out to make room for an additional item that is added to the bottom of the list and how much time to takes before the next item is displayed.


```JavaScript 

var first = 0;
var speed = 700;
var pause = 3000;
``` 

Next, the removeFirst function causes the first item at the top of the list to fade out at a specified speed. The code uses the first pseudo code operator to choose the first list item. It then uses chaining to string together the animate and fadeout functions to accomplish the fade out. And finally, the code calls the addLast function to add an item to the bottom of the list.

```JavaScript 

function removeFirst() {
    first = $('ul#listticker li:first').html();
    $('ul#listticker li:first')
    .animate({ opacity: 0 }, speed)
    .fadeOut('slow', function () { $(this).remove(); });
    addLast(first);
    } 
```
 

The code next sets the i counter that is appended to each line item as it is added at the bottom of the list. The addLast function then adds an item to the bottom of the list. The function uses the intervals that were specified previously to affect the speed at which the items fade into view. The time it takes before the item is displayed is also set by using the setInterval function.

```JavaScript 

var i = 9;

function addLast(first) {
    last = '<li><p><img src="../images/news.gif" /><a href="#">News Headline ' + (i++) + '</a></p></li>';

    first + '';
    $('ul#listticker').append(last)
    $('ul#listticker li:last')
    .animate({ opacity: 1 }, speed)
    .fadeIn('slow')
}
interval = setInterval(removeFirst, pause);
``` 

All other files are automatically provided by the Visual Studio project template for add-ins for Office, and they have not been modified in the development of this sample app.

*Configure the sample*

To configure the sample, open the News_Ticker.sln file with Visual Studio. No other configuration is necessary.

*Build the sample*

To build the sample, choose Ctrl+Shift+B, or on the Build menu, select Build Solution.

*Run and test the sample*

To run the sample, choose the F5 key. After the task pane is displayed in Excel 2013, notice that there are eight lines of text displayed preceded by an image. After an interval of time passes, the top line item starts to fade out. After the top line item disappears, the entire list moves up one position to take the place of the first item. After another interval of time passes, a line item fades in at the bottom of the list. The process then repeats itself.

*Troubleshooting*

If the app fails to install, ensure that the XML in your AnimatedDashboard.xml manifest file parses correctly. Also look for any errors in the JavaScript code that could keep the list from being displayed or the list items from properly fading in or fading out. For example, you may have forgotten to end a statement with a semicolon, or you may have misspelled a method name or keyword. If the components in the task pane do not look as you think they should, check the CSS styles to ensure that you didn't forget a colon between the style and its value, or leave off a semicolon at the end of a style statement.

*Change log*

* First release: April 29, 2013.
* GitHub releas: August 20, 2015

*Related content*

* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Build apps for Office](http://msdn.microsoft.com/library/jj220060.aspx)
* [HTML Tutorial](http://www.w3schools.com/html/)
* [What is jQuery?](http://jquery.com/)
* [CSS Introduction](http://www.w3schools.com/css/css_intro.asp)


