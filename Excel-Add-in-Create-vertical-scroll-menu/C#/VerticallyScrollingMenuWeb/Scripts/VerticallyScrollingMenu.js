/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// Thanks to Kevin Liew (http://www.queness.com/post/256/vertical-scroll-menu-with-jquery-tutorial) for this code.
Office.initialize = function (reason) {
    $(document).ready(function () {

        //Background color, mouseover and mouseout
        var colorOver = '#31b8da';
        var colorOut = '#1f1f1f';

        //Padding, mouseover
        var padLeft = '20px';
        var padRight = '20px'

        //Default Padding
        var defpadLeft = $('#menu li a').css('paddingLeft');
        var defpadRight = $('#menu li a').css('paddingRight');

        //Animate the LI on mouse over, mouse out
        $('#menu li').click(function () {
            //Make LI clickable
            window.location = $(this).find('a').attr('href');

        }).mouseover(function () {

            //Mouse over LI and look for A element for transition
            $(this).find('a')
            .animate({ paddingLeft: padLeft, paddingRight: padRight }, { queue: false, duration: 100 })
            .animate({ backgroundColor: colorOver }, { queue: false, duration: 200 });

        }).mouseout(function () {

            //Mouse oout LI and look for A element and discard the mouse over transition
            $(this).find('a')
            .animate({ paddingLeft: defpadLeft, paddingRight: defpadRight }, { queue: false, duration: 100 })
            .animate({ backgroundColor: colorOut }, { queue: false, duration: 200 });
        });

        //Scroll the menu on mouse move above the #sidebar layer
        $('#sidebar').mousemove(function (e) {

            //Sidebar Offset, Top value
            var s_top = parseInt($('#sidebar').offset().top);

            //Sidebar Offset, Bottom value
            var s_bottom = parseInt($('#sidebar').height() + s_top);

            //Roughly calculate the height of the menu by multiply height of a single LI with the total of LIs
            var mheight = parseInt($('#menu li').height() * $('#menu li').length);

            //Calculate the top value
            var top_value = Math.round(((s_top - e.pageY) / 100) * mheight / 2)

            //Animate the #menu by changing the top value
            $('#menu').animate({ top: top_value }, { queue: false, duration: 500 });
        });

    });
};
// *********************************************************
//
// Excel-Add-in-Create-vertical-scroll-menu, https://github.com/OfficeDev/Excel-Add-in-Create-vertical-scroll-menu
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