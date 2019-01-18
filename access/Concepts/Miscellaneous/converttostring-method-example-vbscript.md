---
title: ConvertToString method example (VBScript)
ROBOTS: INDEX
ms.prod: access
ms.assetid: e2315ef1-41ff-22b6-2417-6eba1f5f06d7
ms.date: 06/08/2017
localization_priority: Normal
---


# ConvertToString method example (VBScript)

**Applies to:** Access 2013 | Access 2016

The following example shows how to convert a **Recordset** into a MIME-encoded string using the **RDSServer.DataFactory** **ConvertToString** method. It then shows how the string can be converted back into a **Recordset**. Cut and paste the following code to Notepad or another text editor and save it as **ConvertToString.htm**.

```vb
<!-- BeginConvertToStringVBS --><HTML>
<HEAD><TITLE>ConvertToString Example</TITLE><HEAD><BODY> 
<SCRIPT LANGUAGE=VBSCRIPT>Sub ConvertToStringX()
Dim objRs, objDF, strServer, vStringConst adcExecSync = 1
Const adcFetchUpFront = 1 
' Replace value below with your server name to use without ASP.strServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>"> 
Set objDF = RDS1.CreateObject("RDSServer.DataFactory", strServer)Set objRs = objDF.Query(txtConnect.Value,txtQueryRecordset.Value) 
' convert Recordset to MIME encoded stringvString = objDF.ConvertToString(objRs) 
' display MIME string for demo purposestxtRS.value = vString 
' convert MIME string back to useable ADO Recordset' using RDS.DataControl
RDC1.SQL = vString 
RDC1.ExecuteOptions = adcExecSyncRDC1.FetchOptions = adcFetchUpFront
RDC1.Refresh 
MsgBox "RecordCount = " & RDC1.Recordset.RecordCountEnd Sub
</SCRIPT> 
Connect String:<INPUT TYPE=Text NAME=txtConnect SIZE=50
VALUE="Provider=sqloledb;Initial Catalog=pubs;Integrated Security='SSPI';"><BR> 
Query:<INPUT TYPE=Text NAME=txtQueryRecordset SIZE=50
VALUE="select * from authors"><BR> 
<INPUT TYPE=Button VALUE="ConvertToString" OnClick="ConvertToStringX()"><BR> 
MIME Encoded RS: <BR><TEXTAREA NAME=txtRS ROWS=15 COLS=50 WRAP=virtual></TEXTAREA> 
<!-- RDS.DataSpace ID RDS1 --><OBJECT ID="RDS1" WIDTH=1 HEIGHT=1
CLASSID="CLSID:BD96C556-65A3-11D0-983A-00C04FC29E36"></OBJECT> 
<!-- RDS.DataControl ID RDC1 --><OBJECT ID="RDC1" WIDTH=1 HEIGHT=1
CLASSID="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"></OBJECT>
</BODY></HTML>
<!-- EndConvertToStringVBS -->

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]