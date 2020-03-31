---
title: DataControl object example (VBScript)
ms.prod: access
ms.assetid: 8e7b613c-6dfc-5c47-5f96-67b7c18d294f
ms.date: 06/08/2019
localization_priority: Normal
---


# DataControl object example (VBScript)

**Applies to:** Access 2013 | Access 2016

The following code shows how to set the [RDS.DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) parameters at design time and bind them to a data-aware control. Cut and paste this code between the `<Body>` and `</Body>` tags in a normal HTML document and name it **DataControlDesignVBS.asp**. ASP script will identify your server.


```vb

<!-- BeginDataControlDesignVBS --><%@ Language=VBScript %>
<HTML><HEAD>
<META name="VI60_DefaultClientScript" content=VBScript><META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>RDS DataControl</title> 
<%' local style sheet used for display%><STYLE>
<!--BODY {
font-family: 'Verdana','Arial','Helvetica',sans-serif;BACKGROUND-COLOR:white;
COLOR:black;}
.thead {background-color: #008080;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.thead2 {background-color: #800000;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.tbody {text-align: center;
background-color: #f7efde;font-family: 'Verdana','Arial','Helvetica',sans-serif;
font-size: x-small;}
--></STYLE>
</HEAD> 
<BODY><H2>RDS API Code Examples</H2>
<HR><H3>Remote Data Service</H3>
<TABLE DATASRC=#RDS><TBODY>
<TR><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD></TR>
</TBODY></TABLE>
<!-- Remote Data Service with Parameters set at Design Time --> 
<OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"ID=RDS>
<PARAM NAME="SQL" VALUE="Select * from Employees for browse"><PARAM NAME="SERVER" VALUE="https://<%=Request.ServerVariables("SERVER_NAME")%>">
<PARAM NAME="CONNECT" VALUE="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'"></OBJECT> 
</BODY></HTML>
<!-- EndDataControlDesignVBS -->
```

The following example shows how to set the necessary parameters of **RDS.DataControl** at run time. To test this example, cut and paste this code between the `<Body>` and `</Body>` tags in a normal HTML document and name it **DataControlRuntimeVBS.asp**. ASP script will identify your server.


```vb

<!-- BeginDataControlRuntimeVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Data Control Object Example (VBScript)</title> 
<%' local style sheet used for display%><style>
<!--body {
font-family: 'Verdana','Arial','Helvetica',sans-serif;BACKGROUND-COLOR:white;
COLOR:black;}
.thead {background-color: #008080;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.thead2 {background-color: #800000;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.tbody {text-align: center;
background-color: #f7efde;font-family: 'Verdana','Arial','Helvetica',sans-serif;
font-size: x-small;}
--></style>
</head> 
<body><h1>Data Control Object Example (VBScript)</h1> 
<H2>RDS API Code Examples</H2><HR>
<H3>Remote Data Service Run Time</H3> 
<TABLE DATASRC=#RDS><TBODY>
<TR><TD><SPAN DATAFLD="au_lname"></SPAN></TD>
<TD><SPAN DATAFLD="au_fname"></SPAN></TD></TR>
</TBODY></TABLE>
<% ' RDS.DataControl with no parameters set at design time %><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS HEIGHT=1 WIDTH=1></OBJECT> 
<FORM name="frmInput"><HR>
<Input Size="70" Name="txtServer" Value="https://<%=Request.ServerVariables("SERVER_NAME")%>"><BR><Input Size="100" Name="txtConnect" Value="Provider='sqloledb';Data Source=<%=Request.ServerVariables("SERVER_NAME")%>;Initial Catalog='Pubs';Integrated Security='SSPI';">
<BR><Input Size="70" Name="txtSQL" Value="Select * from Authors">
<HR><INPUT TYPE="BUTTON" NAME="Run" VALUE="Run"><BR>
<H4>Show grid with these values or change them to see data from another ODBC data source on your server</H4></FORM> 
<Script Language="VBScript"> 
' Set parameters of RDS.DataControl at Run TimeSub Run_OnClick 
RDS.Server = document.frmInput.txtServer.ValueRDS.Connect = document.frmInput.txtConnect.Value
RDS.SQL = document.frmInput.txtSQL.Value 
RDS.Refresh 
End Sub 
</Script> 
</body></html>
<!-- EndDataControlRuntimeVBS -->
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]