<div align="center">

## Client\-Side Form Field Validation with VBScript and ASP


</div>

### Description

Form field validation is one of the most important functions in making successful web-based applications, but server-side validation is not always the answer!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Glenn Cook](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/glenn-cook.md)
**Level**          |Beginner
**User Rating**    |3.6 (25 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/glenn-cook-client-side-form-field-validation-with-vbscript-and-asp__4-6172/archive/master.zip)





### Source Code


<p><strong><font face="Verdana">Client-Side Form Field Validation with VBScript
and ASP</font></strong></p>
<p><font face="Verdana"><small>Form field validation is one of the most
important functions in making successful web-based applications. It protects
your database from errors, filters bad data, and forces the user to think twice
before submitting data with typos. For some reason though, documentation in many
VBScript and JavaScript books on the market tend to gloss over this issue. I
think it might be that field validation scripting is somewhere between learning
the basics but not complicated enough for the advanced literature. The ASP books
cover server-side validation (Remember? It</small><small>'s a server-side
scripting language.) but this is not always an ideal approach to validating
data. The problems with server-side validation are that it bogs down your server</small><small>'s
resources when the user could handle the routines locally, and it offers a
slower response time for the user when it actually finds errors. This tutorial
is designed to take the load off the servers for a while and get the beginning
weblication programmer closer to the intermediate level and beyond. Have fun!</small></font></p>
<p><small><strong><font face="Verdana">The goal of the script:</font></strong></small></p>
<p><small><font face="Verdana">When I click the submit button I want my script
to check the field data in my form and either:</font></small></p>
<blockquote>
 <ol>
  <li><small><font face="Verdana">Submit the data if it all looks good</font></small>
  <li><small><font face="Verdana">Show an alert box and do not send the data.</font></small></li>
 </ol>
</blockquote>
<p><strong><small><font face="Verdana">You&nbsp; need a few tools:</font></small></strong></p>
<blockquote>
 <ul>
  <li><small><font face="Verdana">The OnSubmit event.</font></small>
  <li><small><font face="Verdana">A form name in your HTML code (e.g. </font></small><font face="Courier New">&lt;</font><font color="#7f007f" face="Courier New">form
   </font><font color="#ff0000" face="Courier New">METHOD=&quot;</font><font color="#0000ff" face="Courier New">POST</font><font color="#ff0000" face="Courier New">&quot;
   ACTION=&quot;vbsample.asp&quot; name=&quot;</font><font color="#0000ff" face="Courier New">MyForm</font><font color="#ff0000" face="Courier New">&quot;</font><font face="Courier New">&gt;)</font>
  <li><small><font face="Verdana">A Submit Button: (e.g.</font></small> <font face="Courier New">&lt;<font color="#7f007f">input
   </font><font color="#ff0000">TYPE=&quot;</font><font color="#0000ff">submit</font><font color="#ff0000">&quot;
   VALUE=&quot;</font><font color="#0000ff">Submit Info!</font><font color="#ff0000">&quot;</font><font color="#ff0000" size="1">
   </font><font color="#ff0000">name=&quot;</font><font color="#0000ff">submit</font><font color="#ff0000">&quot;</font>&gt;</font>
  <li><font face="Verdana"><small>An Input Box with name: (e.g.</small><big> </big></font><font face="Courier New"><font color="#000000" size="3">&lt;</font><font color="#800080" size="3">input
   </font><font color="#ff0000" size="3">type=&quot;</font><font color="#0000ff" size="3">text</font><font color="#ff0000" size="3">&quot;
   name=&quot;</font><font color="#0000ff" size="3">MyBox</font><font color="#ff0000" size="3">&quot;
   size=&quot;</font><font color="#0000ff" size="3">10</font><font color="#ff0000" size="3">&quot;</font><font color="#000000" size="3">&gt;</font></font><font color="#000000" size="1">)</font>
  <li><small><font face="Verdana">Some &quot;If..Then&quot; statements</font></small>
  <li><small><font face="Verdana">An &quot;Alert Box&quot;</font></small>
  <li><small><font face="Verdana">The secret. (I love suspense.)</font></small></li>
 </ul>
</blockquote>
<p><strong><font face="Verdana">Let's see it work before we get into the guts of
it!</font></strong>
<ol>
 <li><small><font face="Verdana">Type your name and submit!</font></small>
 <li><small><font face="Verdana">Enter nothing and submit!</font></small>
 <li><small><font face="Verdana">Enter one character and submit!</font></small></li>
</ol>
<div align="center">
 <form action="http://www.aspalliance.com/glenncook/vbsample.asp" method="post" name="MyForm">
  <p><input name="MyBox" size="20"><br>
  <input name="s1" type="submit" value="Submit Info!"><input name="B2" type="reset" value="Clear"><br>
  &nbsp;<img src="http://www.aspalliance.com/glenncook/images/kidasper.jpg" style="LEFT: 160px; TOP: 759px" width="128" height="125"></p>
 </form>
</div>
<font face="Courier New" size="1">
<div align="left">
 <table border="1" cellPadding="4" width="100%">
  <tbody>
   <tr>
    <td align="left" vAlign="top" width="50%"></font><font color="#0000ff" face="Courier New">&lt;SCRIPT
    LANGUAGE=&quot;VBScript&quot;&gt;</font><font face="Courier New" size="2">
    <p><font color="#008040">&lt;-- Option Explicit</font></p>
    <p><font color="#008040">dim validation</font></p>
    <p><font color="#008040">dim header</font></p>
    <p><font color="#008040">header = aspalliance.com</font></font></p>
   </td>
   <td align="left" vAlign="top" width="50%"><font face="Verdana"><small>The
    first thing you want to do is tell the browser that the code you are
    about to run is VBScript. &nbsp; The next thing you want to do is make
    some variables that we are going to use in our routine.</small><br>
    <small>The &quot;<strong>validation</strong>&quot; variable is going to
    be used to help us determine whether we should send the form or not.&nbsp;
    It will can be equal to either &quot;True&quot; or &quot;False&quot; (By
    the way, that's pretty much the secret ingredient!</small></font><small><font face="Verdana" size="1">)</font></small><font face="Verdana"><br>
    <small>The &quot;<strong>Header</strong>&quot; variable I will use to
    hold the string information for a MsgBox property.</small></font></td>
  </tr>
  <tr>
   <td align="left" vAlign="top" width="50%"><font color="#008040" face="Courier New" size="2">Function
    MyForm_OnSubmit</font><font face="Courier New" size="1">
    <p></font><font color="#008040" face="Courier New" size="2">validation =
    True</font></p>
   </td>
   <td align="left" vAlign="top" width="50%"><font face="Verdana"><small>The
    next thing you want to do is make a <strong>function</strong> that is
    going to be called when a user clicks the &quot;Submit Info&quot;
    button. The function needs to be bound to your form name and the <strong>OnSubmit</strong>
    event so VB knows what it is that you are trying to submit.&nbsp; I have
    named my form....&quot;<strong>MyForm</strong>&quot;</small><br>
    <small>I have also made my &quot;validation&quot; variable equal to
    &quot;<strong>True</strong>.&quot; &nbsp; I will follow this with some
    &quot;If...Then&quot; statements that will try to change the variable to
    &quot;False.&quot;</small></font></td>
  </tr>
  <tr>
   <td align="left" vAlign="top" width="50%"><font color="#008040" face="Courier New"><small>If
    Len(Document.MyForm.MyBox.Value) &gt; 2 Then</small></font>
    <p><font color="#008040" face="Courier New"><small>MsgBox &quot;You have
    entered too many characters!&nbsp; You need fewer characters before I
    will submit this form&quot;,8, Header</small></font></p>
    <p><font color="#008040" face="Courier New"><small>validation = False</small></font></p>
    <p><font color="#008040" face="Courier New"><small>End If</small></font></p>
   </td>
   <td align="left" vAlign="top" width="50%"><font face="Verdana"><small>My
    first &quot;If...Then&quot; basically says that <strong>If </strong>the <strong>Len</strong>gth
    of the<strong> value</strong> of the contents of <strong>Mybox</strong>
    in the form <strong>Myform</strong> within the current <strong>document</strong>
    is less than five characters <strong>Then </strong>show the Message Box(<strong>MsgBox</strong>).
    The properties of the message box include the message string, then the
    type of message box(<strong>8</strong>), and I add the <strong>header</strong>
    message.&nbsp; The next thing I do is make our validation variable equal
    to <strong>False</strong>. You'll see why in a second.</small><br>
    <small>The last thing I do is <strong>End</strong> this <strong>If </strong>statement.</small></font></td>
  </tr>
  <tr>
   <td align="left" vAlign="top" width="50%"><font color="#008040" face="Courier New"><small>If
    (Document.MyForm.MyBox.Value) = &quot;&quot; Then</small></font>
    <p><font color="#008040" face="Courier New"><small>MsgBox &quot;You have
    forgotten to fill in the input box!&nbsp; Why would you want to submit
    nothing in a one field form? C'mon, give me &nbsp; something to work
    with!&quot;,8, Header</small></font></p>
    <p><font color="#008040" face="Courier New"><small>validation = False</small></font></p>
    <p><font color="#008040" face="Courier New"><small>End If</small></font></p>
   </td>
   <td align="left" vAlign="top" width="50%"><small><font face="Verdana">In
    this &quot;If..Then&quot; I make sure the user has at least entered
    something into the text input box.&nbsp; If they haven't they will get a
    message box and once again the validation variable will be made equal to
    false.</font></small></td>
  </tr>
  <tr>
   <td align="left" vAlign="top" width="50%"><font color="#008040" face="Courier New"><small>If
    validation = True Then</small></font>
    <p><font color="#008040" face="Courier New"><small>MyForm_OnSubmit =
    True</small></font></p>
    <p><font color="#008040" face="Courier New"><small>Else</small></font></p>
    <p><font color="#008040" face="Courier New"><small>MyForm_OnSubmit =
    False</small></font></p>
    <p><font color="#008040" face="Courier New"><small>End If</small></font></p>
    <p><font color="#008040" face="Courier New"><small>End Function</small></font></p>
    <p><font color="#0000ff" face="Courier New">&lt;/SCRIPT&gt;</font></p>
   </td>
   <td align="left" vAlign="top" width="50%"><font face="Verdana"><small>The
    last part of this routine is the secret.</small><br>
    <small>All I do is test to see if any &quot;If..Then&quot; statement has
    made my validation variable equal to something other than
    &quot;True&quot;.&nbsp; <strong>If</strong> the <strong>validation</strong>
    variable equals &quot;<strong>True</strong>&quot; then <strong>MyForm_OnSubmit</strong>
    also equals &quot;<strong>True</strong>.&quot;&nbsp; If the variable
    equals something <strong>else</strong> other than &quot;True&quot; Then <strong>MyForm_OnSubmit</strong>
    will equal &quot;<strong>False</strong>.&quot; &nbsp; If the Function
    equals &quot;False&quot; then the OnSubmit event will not execute and
    the user is prompted to fix their mistake. Otherwise, the OnSubmit event
    fires and the form data is <strong>post</strong>ed to the page that will
    process the submitted form. Simple!</small></font>
    <p><font face="Verdana"><small>Finally, <strong>End</strong> your <strong>If</strong>,
    tell VBScript to end the <strong>Function</strong>, and close the script
    with the <strong>SCRIPT</strong> tag.</small></font></p>
   </td>
  </tr>
 </tbody>
</table>
</div>
<p><strong><font face="Verdana">Now granted, this is is a very simple example
and in a real-world application you might have ten to twenty form fields to
validate. &nbsp; Although there is potential for your script becoming quite
complicated you can use some subroutines that do some work for you.&nbsp;
Example:</font></strong></p>
<p><font color="#008040" face="Courier New">Function MyForm_OnSubmit</font></p>
<p><font color="#008040" face="Courier New">Call
Check(Document.MyForm.Company.Value,&nbsp;&nbsp;&nbsp;&nbsp; &quot;Please enter
a company name.&quot;)<big><br>
</big>Call Check(Document.MyForm.Name.Value,&nbsp;&nbsp;&nbsp;&nbsp;
&quot;Please enter a name.&quot;)</font></p>
<p><font color="#008040" face="Courier New">Sub Check(ByVal FieldValue, ByVal
message)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; If FieldValue = &quot;&quot; Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; MsgBox
message, 8, Header<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; validation =
False<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End Sub</font><big><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</big></p>

