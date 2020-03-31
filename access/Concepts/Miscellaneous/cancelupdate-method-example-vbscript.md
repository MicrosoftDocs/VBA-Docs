---
title: CancelUpdate method example (VBScript)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 5442db02-960c-6d11-e0f3-0726c8d529bd
ms.date: 06/08/2019
localization_priority: Normal
---


# CancelUpdate method example (VBScript)

**Applies to:** Access 2013 | Access 2016

To test this example, cut and paste this code between the `<Body>` and `</Body>` tags in a normal HTML document and name it **CancelUpdateVBS.asp**. ASP script will identify your Internet server. You will need to edit the name of the server to reflect your own setup. Simply change the value in the connect string from MyServer to the name of your SQL Server installation.

```vb
<!-- BeginCancelUpdateVBS --><%@Language=VBScript%> 
<%'Option Explicit%><% 'use the following META tag instead of adovbs.inc%>
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --><HTML>
<HEAD><META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE></HEAD>
<BODY><CENTER>
<H1>Remote Data Service</H1><H2>SubmitChanges and CancelUpdate Methods</H2> 
<% ' to integrate/test this code replace the Server property value and' the Data Source value in the Connect property with appropriate values%> 
<HR><OBJECT ID=RDS classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" HEIGHT=1 WIDTH=1></OBJECT>
<SCRIPT Language="VBScript"> 
'set RDS properties for control just created 
RDS.Server = "https://<%=Request.ServerVariables("SERVER_NAME")%>"RDS.SQL = "Select * from Employees"
RDS.Connect = "Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind';"RDS.Refresh
</SCRIPT> 
<TABLE DATASRC=#RDS><THEAD>
<TR ID="ColHeaders"><TH>ID</TH>
<TH>FName</TH><TH>LName</TH>
<TH>Title</TH><TH>Hire Date</TH>
<TH>Birth Date</TH><TH>Extension</TH>
<TH>Home Phone</TH></TR>
</THEAD><TBODY>
<TR><TD> <INPUT DATAFLD="EmployeeID" size=4> </TD>
<TD> <INPUT DATAFLD="FirstName" size=10> </TD><TD> <INPUT DATAFLD="LastName" size=10> </TD>
<TD> <INPUT DATAFLD="Title" size=10> </TD><TD> <INPUT DATAFLD="HireDate" size=10> </TD>
<TD> <INPUT DATAFLD="BirthDate" size=10> </TD><TD> <INPUT DATAFLD="Extension" size=10> </TD>
<TD> <INPUT DATAFLD="HomePhone" size=8> </TD></TR>
</TBODY></TABLE>
<HR><INPUT TYPE=button NAME="SubmitChange" VALUE="Submit Changes">
<INPUT TYPE=button NAME="CancelChange" VALUE="Cancel Update"><BR>
<H4>Alter a current entry on the grid. Move off that Row. <BR>Submit the Changes to your DBMS or cancel the updates. </H4>
</CENTER><SCRIPT Language="VBScript"> 
Sub SubmitChange_OnClick 
msgbox "Changes will be made"RDS.SubmitChanges
RDS.Refresh 
End Sub 
Sub CancelChange_OnClick 
msgbox "Changes will be cancelled"RDS.CancelUpdate
RDS.Refresh 
End Sub-->
</SCRIPT> 
 
</BODY></HTML>
<!-- EndCancelUpdateVBS -->
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]