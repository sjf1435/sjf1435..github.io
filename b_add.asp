<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level1.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_conn_bargain_STRING
  MM_editTable = "t_bargain"
  MM_editRedirectUrl = "bargain.asp"
  MM_fieldsStr  = "b_tid|value|b_num|value|b_name|value|b_company|value|b_date|value|b_datediff|value|b_money|value|b_sign|value|b_check|value|b_aname|value|b_content|value"
  MM_columnsStr = "b_tid|none,none,NULL|b_num|',none,''|b_name|',none,''|b_company|',none,''|b_date|',none,''|b_datediff|none,none,NULL|b_money|',none,''|b_sign|',none,''|b_check|',none,''|b_aname|',none,''|b_content|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
	  if cint(session("mm_level")) = 1 then
	  response.Redirect "b_list.asp"
	  else
      Response.Redirect(MM_editRedirectUrl)
	  end if
    End If
  End If

End If
%>
<%
Dim rsbt
Dim rsbt_numRows

Set rsbt = Server.CreateObject("ADODB.Recordset")
rsbt.ActiveConnection = MM_conn_bargain_STRING
rsbt.Source = "SELECT * FROM t_btype"
rsbt.CursorType = 0
rsbt.CursorLocation = 2
rsbt.LockType = 1
rsbt.Open()

rsbt_numRows = 0
%>
<% 
If rsbt.EOF And rsbt.BOF Then 
response.Write("���������ҵ���")
response.end
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Config/style.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" type="text/css" media="all" href="Config/calendar-win2k-1.css" title="win2k-1" />
<script type="text/javascript" src="Config/calendar.js"></script>
<script type="text/javascript" src="lang/calendar-en.js"></script>
<script type="text/javascript">
<!--
var oldLink = null;
// code to change the active stylesheet
function setActiveStyleSheet(link, title) {
  var i, a, main;
  for(i=0; (a = document.getElementsByTagName("link")[i]); i++) {
    if(a.getAttribute("rel").indexOf("style") != -1 && a.getAttribute("title")) {
      a.disabled = true;
      if(a.getAttribute("title") == title) a.disabled = false;
    }
  }
  if (oldLink) oldLink.style.fontWeight = 'normal';
  oldLink = link;
  link.style.fontWeight = 'bold';
  return false;
}

// This function gets called when the end-user clicks on some date.
function selected(cal, date) {
  cal.sel.value = date; // just update the date in the input field.
  if (cal.sel.id == "sel1" || cal.sel.id == "sel3")
    // if we add this call we close the calendar on single-click.
    // just to exemplify both cases, we are using this only for the 1st
    // and the 3rd field, while 2nd and 4th will still require double-click.
    cal.callCloseHandler();
}

// And this gets called when the end-user clicks on the _selected_ date,
// or clicks on the "Close" button.  It just hides the calendar without
// destroying it.
function closeHandler(cal) {
  cal.hide();                        // hide the calendar
}

// This function shows the calendar under the element having the given id.
// It takes care of catching "mousedown" signals on document and hiding the
// calendar if the click was outside.
function showCalendar(id, format) {
  var el = document.getElementById(id);
  if (calendar != null) {
    // we already have some calendar created
    calendar.hide();                 // so we hide it first.
  } else {
    // first-time call, create the calendar.
    var cal = new Calendar(false, null, selected, closeHandler);
    // uncomment the following line to hide the week numbers
    // cal.weekNumbers = false;
    calendar = cal;                  // remember it in the global var
    cal.setRange(1900, 2070);        // min/max year allowed.
    cal.create();
  }
  calendar.setDateFormat(format);    // set the specified date format
  calendar.parseDate(el.value);      // try to parse the text in field
  calendar.sel = el;                 // inform it what input field we use
  calendar.showAtElement(el);        // show the calendar below it

  return false;
}

var MINUTE = 60 * 1000;
var HOUR = 60 * MINUTE;
var DAY = 24 * HOUR;
var WEEK = 7 * DAY;

// If this handler returns true then the "date" given as
// parameter will be disabled.  In this example we enable
// only days within a range of 10 days from the current
// date.
// You can use the functions date.getFullYear() -- returns the year
// as 4 digit number, date.getMonth() -- returns the month as 0..11,
// and date.getDate() -- returns the date of the month as 1..31, to
// make heavy calculations here.  However, beware that this function
// should be very fast, as it is called for each day in a month when
// the calendar is (re)constructed.
function isDisabled(date) {
  var today = new Date();
  return (Math.abs(date.getTime() - today.getTime()) / DAY) > 10;
}

function flatSelected(cal, date) {
  var el = document.getElementById("preview");
  el.innerHTML = date;
}

function showFlatCalendar() {
  var parent = document.getElementById("display");

  // construct a calendar giving only the "selected" handler.
  var cal = new Calendar(false, null, flatSelected);

  // hide week numbers
  cal.weekNumbers = false;

  // We want some dates to be disabled; see function isDisabled above
  cal.setDisabledHandler(isDisabled);
  cal.setDateFormat("DD, M d");

  // this call must be the last as it might use data initialized above; if
  // we specify a parent, as opposite to the "showCalendar" function above,
  // then we create a flat calendar -- not popup.  Hidden, though, but...
  cal.create(parent);

  // ... we can show it here.
  cal.show();
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<form action="<%=MM_editAction%>" method="POST" name="form1" onSubmit="MM_validateForm('b_name','','R','b_company','','R','sel2','','R','b_datediff','','RisNum','b_money','','RisNum','b_content','','R');return document.MM_returnValue">
  <h1>��ͬ¼��</h1>
  <hr size="1">
  <table width="100%" align="center">
    <tr> 
      <td width="30%" height="30" align="right" nowrap class="bgcolor-left"> ��ͬ����:</td>
      <td height="30" class="bgcolor-right"> <select name="b_tid" class="px12">
          <%
While (NOT rsbt.EOF)
%>
          <option value="<%=(rsbt.Fields.Item("bt_id").Value)%>"><%=(rsbt.Fields.Item("bt_name").Value)%></option>
          <%
  rsbt.MoveNext()
Wend
If (rsbt.CursorType > 0) Then
  rsbt.MoveFirst
Else
  rsbt.Requery
End If
%>
        </select> </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��ͬ��:</td>
      <td height="30" class="bgcolor-right"> <% dim mm_bnum
  mm_bnum = now()
  mm_bnum = replace(mm_bnum,"-","")
  mm_bnum = replace(mm_bnum," ","") 
  mm_bnum = replace(mm_bnum,":","")
  mm_bnum = replace(mm_bnum,"PM","")
  mm_bnum = replace(mm_bnum,"AM","")
  mm_bnum = replace(mm_bnum,"����","")
  mm_bnum = replace(mm_bnum,"����","") %> <%=mm_bnum%> <input name="b_num" type="hidden" id="b_num" value="<%=mm_bnum%>"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��ͬ����:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="b_name" value="" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��չ��:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="b_company" value="" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��ͬ����:</td>
      <td height="30" class="bgcolor-right"><input type="text" name="b_date" id="sel2" size="32"
> <input name="reset" type="reset"
onClick="return showCalendar('sel2', 'dd/mm/y');" value=" ... "></td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��ͬ��:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="b_datediff" value="" size="5">
        �� </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">��ͬ�ܽ��:</td>
      <td height="30" class="bgcolor-right"> <input type="text" name="b_money" value="" size="32"> 
      </td>
    </tr>
    <tr> 
      <td height="30" align="right" nowrap class="bgcolor-left">�Ƿ�ǩԼ:</td>
      <td height="30" class="bgcolor-right"> <input type="radio" name="b_sign" value="��ǩԼ">
        ��ǩԼ 
        <input type="radio" name="b_sign" value="δǩԼ">
        δǩԼ 
        <input name="b_check" type="hidden" id="b_check" value="δ���"> <input name="b_aname" type="hidden" id="b_aname" value="<%=session("mm_name")%>"></td>
    </tr>
    <tr> 
      <td align="right" nowrap class="bgcolor-left">��ͬ����:</td>
      <td class="bgcolor-right"> <textarea name="b_content" cols="80" rows="10" class="px12"></textarea> 
      </td>
    </tr>
    <tr> 
      <td align="right" valign="baseline" nowrap>&nbsp;</td>
      <td valign="baseline"> <input type="submit" value="¼���ͬ"> </td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
<p>&nbsp;</p>

</body>
</html>
<%
rsbt.Close()
Set rsbt = Nothing
%>
