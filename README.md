<div align="center">

## Using LSet to make reading files easier


</div>

### Description

Have you ever needed to read a fixed length text file in Visual Basic? I have written quite a lot of applications where I will receive a fixed length text file from the mainframe.

Tired of using the Left$, Right$ and Mid$ functions to parse each line for the individual elements of the line? Then read on.

Oh, and thanks to Rockford Lhotka's, book on Visual Basic 6 Business Objects for giving me this idea.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-09-18 21:56:04
**By**             |[Jerry Barnett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jerry-barnett.md)
**Level**          |Intermediate
**User Rating**    |4.7 (103 globes from 22 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD99629182000\.zip](https://github.com/Planet-Source-Code/jerry-barnett-using-lset-to-make-reading-files-easier__1-10755/archive/master.zip)





### Source Code

<p>Using the LSET keyword you can use a user defined type (UDT)
to automatically parse a line from a text file into the
individual elements......in ONE LINE OF CODE! Here is how:</p>
<p>First lets take a sample file (call it 'SOMEFILE.TXT'):</p>
<pre>08072000Jerry M Barnett 0002356A2S56D9</pre>
<p>Ok, the line above represents lets say one of a few hundred
lines. The layout of the file is as follows:</p>
<table border="0" cellpadding="4">
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2"><strong>Field</strong></font></td>
 <td><font size="2"><strong>Type</strong></font></td>
 <td><font size="2"><strong>Position</strong></font></td>
 <td><font size="2"><strong>Remarks</strong></font></td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2">Date</font></td>
 <td><font size="2">MMDDYYYY</font></td>
 <td><font size="2">1-8</font></td>
 <td><font size="2">Format will be MM/DD/YYYY</font></td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2">Name</font></td>
 <td><font size="2">AlphaNumeric</font></td>
 <td><font size="2">9-29</font></td>
 <td><font size="2">Padded with space</font></td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2">Amount</font></td>
 <td><font size="2">Numeric</font></td>
 <td><font size="2">30-36</font></td>
 <td><font size="2">9(5)v99</font></td>
 <td><font size="2"></font> </td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2">Code</font></td>
 <td><font size="2">Alpha</font></td>
 <td><font size="2">37</font></td>
 <td><font size="2">A for accepted,</font></td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2"></font> </td>
 <td><font size="2"></font> </td>
 <td><font size="2"></font> </td>
 <td><font size="2">R for rejected</font></td>
 </tr>
 <tr>
 <td><font size="2"></font> </td>
 <td><font size="2">Account</font></td>
 <td><font size="2">AlphaNumeric</font></td>
 <td><font size="2">38-43</font></td>
 <td><font size="2"></font> </td>
 </tr>
</table>
<p>First thing is to create a user defined type representing the
layout of the file. (Note - this should be place in the Module
level of a program.)</p>
<table border="0">
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Type</font></td>
 <td><font size="2" face="Courier New">udtInput</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">MyDate</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">As String * 8</font></td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Name</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">As String * 21</font></td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Amount</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">As String * 7</font></td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Code</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">As String * 1</font></td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Account</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">As String * 6</font></td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">End </font></td>
 <td><font size="2" face="Courier New">Type</font></td>
 </tr>
</table>
<p><strong>Note</strong> - All types are strings reguardless of
the type in the file. This will become clearer later. Next,
create a user defined type to represent the total length of the
line (I will explain why later)</p>
<table border="0">
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">Type</font></td>
 <td><font size="2" face="Courier New">udtLine</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">strBuff</font></td>
 <td><font size="2" face="Courier New">As String * 43</font></td>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New"></font> </td>
 </tr>
 <tr>
 <td><font size="2" face="Courier New"></font> </td>
 <td><font size="2" face="Courier New">End</font></td>
 <td><font size="2" face="Courier New">Type</font></td>
 <td><font size="2"></font> </td>
 <td><font size="2"></font> </td>
 <td><font size="2"></font> </td>
 </tr>
</table>
<p><strong>Now in your program you can do the following:</strong></p>
<pre>Sub Main()</pre>
<pre><font color="#008000">' File Number</font></pre>
<pre>Dim iMyFile As Long</pre>
<pre>Dim sLine As udtLine</pre>
<pre>Dim sInput As udtInput</pre>
<pre><font color="#008000">' Open your file for reading</font></pre>
<pre>Open App.Path & "\SOMEFILE.TXT" for Input Access Read as #iMyFile</pre>
<pre><font color="#008080">' Read the line into the udtLine UDT strBuff element</font></pre>
<pre><font color="#008080">' This needs to be done, because you can't place a string</font></pre>
<pre><font color="#008080">' directly into the UDT or you will get a type missmatch errorLine </font></pre>
<pre>Input #iMyFile, sLine.strBuff</pre>
<pre><font color="#008080">' The buffer (strBuff) represents the entire line</font></pre>
<pre><font color="#008080">' Now copy the udtLine UDT (sLine) into the udtInput</font></pre>
<pre><font color="#008080">' UDT (sInput) *** ONE LINE OF CODE! ***</font></pre>
<pre>LSet sInput = sLine</pre>
<pre><font color="#008080">' Wolla! You can now access each element of the sInput UDT!</font></pre>
<pre>Debug.Print sInput.Name</pre>
<pre><font color="#008080">' Will print: Jerry M Barnett</font></pre>
<pre><font color="#008080">' (with eight trailing spaces)</font></pre>
<pre><font color="#008080">' Convert numeric String Amount value to a Long value with 2 decimal places</font></pre>
<pre>Debug.Print CDbl(Val(sInput.Amount)/100)</pre>
<pre><font color="#008080">' Would print: 23.56</font></pre>
<pre>Debug.Print MyDateFunction(sInput.MyDate)</pre>
<pre><font color="#008080">' Will print: 08072000 in any format you wish</font></pre>
<pre><font color="#008080">' Note - The MyDateFunction is a function you define</font></pre>
<pre><font color="#008080">' to parse the date string to the proper format you want.</font></pre>
<pre>End Sub</pre>
<p><font color="#000000"><strong>That's it! Hope you find this of
use!</strong></font></p>
<p><font color="#000000"><strong>Note this can also be used (with
modification to' WRITE a flat file out.)</strong></font></p>

