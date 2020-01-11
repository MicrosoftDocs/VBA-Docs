---
title: URL property example (VBScript)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 667f3927-e5fa-4cc9-b341-027177d1d2d8
ms.date: 06/08/2017
localization_priority: Normal
---


# URL property example (VBScript)

**Applies to:** Access 2013 | Access 2016

The following code demonstrates how to set the **URL** property on the client side to specify an .asp file that in turn handles the submission of changes to the data source.

```vb
<!-- BeginURLClientVBS --> 
<%@ Language=VBScript %> 
<html>
<head> 
<meta name="VI60_DefaultClientScript" content=VBScript> 
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"> 
<title>URL Property Example (VBScript)</title> 
<style> 
<!-- 
body { 
font-family: 'Verdana','Arial','Helvetica',sans-serif; 
BACKGROUND-COLOR:white; 
COLOR:black;
} 

.thead { 
background-color: #008080; 
font-family: 'Verdana','Arial','Helvetica',sans-serif; 
font-size: x-small; 
color: white; 
} 

.thead2 { 
background-color: #800000; 
font-family: 'Verdana','Arial','Helvetica',sans-serif; 
font-size: x-small; 
color: white; 
} 

.tbody { 
text-align: center; 
background-color: #f7efde; 
font-family: 'Verdana','Arial','Helvetica',sans-serif; 
font-size: x-small; 
}
--> 

</style> 
</head> 
<body onload=Getdata()>
<h1>URL Property Example (VBScript)</h1> 
<OBJECT classid=clsid:BD96C556-65A3-11D0-983A-00C04FC29E33 height=1 id=ADC width=1>
</OBJECT> 

<table datasrc="#ADC" align="center"> 
<thead> 
<tr id="ColHeaders" class="thead2"> 
<th>FirstName</th> 
<th>LastName</th>
<th>Extension</th> 
</tr> 
</thead> 
<tbody class="tbody"> 
<tr>
<td><input datafld="FirstName" size=15> </td> 
<td><input datafld="LastName" size=25> </td> 
<td><input datafld="Extension" size=15> </td>
</tr> 
</tbody> 
</table>

<script Language="VBScript"> 

Sub Getdata()

msgbox "getdata" 

ADC.URL = "https://<%=Request.ServerVariables("SERVER_NAME")%>/URLServerVBS.asp" 

ADC.Refresh 

End Sub
</script> 
</body> 
</html> 
<!-- EndURLClientVBS -->
```

The server-side code that exists in **URLServerVBS.asp** submits the updated **Recordset** to the data source.



```vb
<!-- BeginURLServerVBS --> 
<%@ Language=VBScript %> 
<% 
' XML output req's 
Response.ContentType = "text/xml" 

const adPersistXML = 1 

' recordset vars 
Dim strSQL, rsEmployees 
Dim strCnxn, Cnxn 

strCnxn = "Provider='sqloledb';Data Source=" & _ 
Request.ServerVariables("SERVER_NAME") & ";" & _ 
"Integrated Security='SSPI';Initial Catalog='Northwind';" 

Set Cnxn = Server.CreateObject("ADODB.Connection") 

Set rsEmployees = Server.CreateObject("ADODB.Recordset") 

strSQL = "SELECT FirstName, LastName, Extension FROM Employees" 

Cnxn.Open strCnxn 

rsEmployees.Open strSQL, Cnxn 

' output as XML 
rsEmployees.Save Response, adPersistXML 

' Clean up 
rsEmployees.Close 
Cnxn.Close
Set rsEmployees = Nothing 
Set Cnxn = Nothing
%> 
<!-- EndURLServerVBS -->
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]