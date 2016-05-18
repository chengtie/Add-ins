/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

// Calculate the monthly payment for the mortgage.
function calculatePayment(loanAmount, loanTerm, percentage) {

    // The following formula is used to calculate the fixed monthly payment (P) 
    // required to fully amortize a loan of L dollars over a term of n months 
    // at a monthly interest rate of c. [If the quoted rate is 6%, for example, c is .06/12 or .005]. 
    //P = L[c(1 + c)n]/[(1 + c)n - 1]

    var monthlyPayment; // = 1978.85;
    var monthlyInterestRate = (percentage / 100) / 12;

    monthlyPayment = loanAmount * ((monthlyInterestRate * Math.pow(1 + monthlyInterestRate, loanTerm)) / (Math.pow(1 + monthlyInterestRate, loanTerm) - 1));

    monthlyPayment = monthlyPayment.toFixed(2);
    return monthlyPayment
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