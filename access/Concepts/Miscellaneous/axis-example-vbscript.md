---
title: Axis example (VBScript)
ms.prod: access
ms.assetid: 7dd2c9ee-4f67-091a-fd7f-f914016e04b4
ms.date: 06/08/2019
localization_priority: Normal
---


# Axis example (VBScript)

**Applies to:** Access 2013 | Access 2016

This Active Server Page displays OLAP data from an MDX Query string and writes the resulting cellset to an HTML table structure.

```vb
<%@ Language=VBScript %> 
<% 
'************************************************************************ 
'*** Active Server Page displays OLAP data from default 
'*** MDX Query string and writes resulting cell set to HTML table 
'*** structure. 
'************************************************************************ 
Response.Buffer=True 
Response.Expires=0 
%> 
<HTML> 
<HEAD> 
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"> 
</HEAD> 
<BODY bgcolor=Ivory> 
<FONT FACE=Verdana> 
 
<% 
 
Dim cat,cst,i,j,strSource,csw,intDC0,intDC1,intPC0,intPC1 
 
'************************************************************************ 
'*** Set Connection Objects for Multidimensional Catalog and Cellset 
'************************************************************************ 
Set cat = Server.CreateObject("ADOMD.Catalog") 
Set cst = Server.CreateObject("ADOMD.CellSet") 
 
'************************************************************************ 
'*** Use default settings of a known OLAP Server 
'*** for Server Name for Connection Set Server Name Session Object 
'*** to default value 
'************************************************************************ 
'*** Must set OLAPServerName to OLAP Server that is 
'*** present on network 
'************************************************************************ 
   OLAPServerName = "Please set to present OLAP Server" 
   cat.ActiveConnection = "Data Source='" & OLAPServerName & _ 
      "';Initial Catalog='FoodMart';Provider='msolap';" 
 
'************************************************************************ 
'*** Use default MDX Query string of a known query that works 
'*** with default server Set MDXQuery Session Object to default value 
'************************************************************************ 
   strSource = strSource & "SELECT " 
   strSource = strSource & "{[Measures].members} ON COLUMNS," 
   strSource = strSource & _ 
      "NON EMPTY [Store].[Store City].members ON ROWS" 
   strSource = strSource & " FROM Sales" 
 
'************************************************************************ 
'*** Set Cell Set Source property to strSource to be passed on cell set '*** open method 
'************************************************************************ 
    cst.Source = strSource 
 
'************************************************************************ 
'*** Set Cell Sets Active connection to use the current Catalogs Active  
'*** connection 
'************************************************************************ 
Set cst.ActiveConnection = cat.ActiveConnection 
 
'************************************************************************ 
'*** Using Open method, Open cell set 
'************************************************************************ 
cst.Open 
 
'************************************************************************ 
'*** Set Dimension Counts minus 1 for Both Axes to intDC0, intDC1 
'*** Set Position Counts minus 1 for Both Axes to intPC0, intPC1 
'************************************************************************ 
intDC0 = cst.Axes(0).DimensionCount-1 
intDC1 = cst.Axes(1).DimensionCount-1 
 
intPC0 = cst.Axes(0).Positions.Count - 1 
intPC1 = cst.Axes(1).Positions.Count - 1 
 
'************************************************************************ 
'*** Create HTML Table structure to hold MDX Query return Record set 
'************************************************************************ 
      Response.Write "<Table width=100% border=1>" 
 
'************************************************************************ 
'*** Loop to create Column header 
'************************************************************************ 
      For h=0 to intDC0 
         Response.Write "<TR>" 
 
'************************************************************************ 
'*** Loop to create spaces in front of Column headers 
'*** to align with Row header 
'************************************************************************ 
         For c=0 to intDC1 
            Response.Write "<TD></TD>" 
         Next 
 
'************************************************************************ 
'*** Iterate through Axes(0) Positions writing member captions to table  
'*** header 
'************************************************************************ 
         For i = 0 To intPC0 
            Response.Write "<TH>" 
            Response.Write "<FONT size=-2>" 
            Response.Write cst.Axes(0).Positions(i).Members(h).Caption 
            Response.Write "</FONT>" 
            Response.Write "</TH>" 
         Next 
         Response.Write "</TR>" 
      Next 
'************************************************************************ 
'*** Use Array values for row header formatting to provide 
'*** spaces under beginning row header titles 
'************************************************************************ 
      For j = 0 To intPC1 
         Response.Write "<TR>" 
         For h=0 to intDC1 
            Response.Write "<TD><B>" 
            Response.Write "<FONT size=-2>" 
            Response.Write cst.Axes(1).Positions(j).Members(h).Caption 
            Response.Write "</FONT>" 
            Response.Write "</B></TD>" 
         Next 
         For k = 0 To intPC0 
            Response.Write "<TD align=right bgcolor=" 
            Response.Write csw 
            Response.Write ">" 
            Response.Write "<FONT size=-2>" 
            Response.Write cst(k, j).FormattedValue 
            Response.Write "</FONT>" 
            Response.Write "</TD>" 
         Next 
         Response.Write "</TR>" 
      Next 
      Response.Write "</Table>" 
 
%> 
</FONT> 
</BODY> 
</HTML> 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]