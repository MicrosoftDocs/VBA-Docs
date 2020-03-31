---
title: CubeDef example (VBScript)
ms.prod: access
ms.assetid: bcd50cc6-fe2b-d47a-a402-cd2ba4662b2d
ms.date: 06/08/2019
localization_priority: Normal
---


# CubeDef example (VBScript)

**Applies to:** Access 2013 | Access 2016

This example displays cube metadata on a web page.

```vb
<%@ Language=VBScript %> 
<% 
Response.Buffer=True 
'Response.Expires=0 
%> 
<html> 
<head> 
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"> 
</head> 
<body> 
 
<% 
Server.ScriptTimeout=360 
Dim cat,cdf,di,hi,le,mem,strServer,strSource,strCubeName 
 
'************************************************************************ 
'*** Set Session Variables 
'************************************************************************ 
Session("CubeName") = Request.Form("strCubeName") 
Session("CatalogName") = Request.Form("strCatalogName") 
Session("ServerName") = Request.Form("strServerName") 
Session("chkDim") = Request.Form("chkDimension") 
Session("chkHier") =  Request.Form("chkHierarchy") 
Session("chkLev") =  Request.Form("chkLevel") 
 
'************************************************************************ 
'*** Create Catalog Object 
'************************************************************************************ 
Set cat = Server.CreateObject("ADOMD.Catalog") 
 
If Len(Session("ServerName")) > 0 Then 
   cat.ActiveConnection = "Data Source='" & Session("ServerName") & "';Initial Catalog='" & Session("CatalogName") & "';Provider='msolap';" 
Else 
'************************************************************************************ 
'*** Must set OLAPServerName to OLAP Server that is 
'*** present on network 
'************************************************************************ 
OLAPServerName = "Please set to present OLAP Server" 
   cat.ActiveConnection = "Data Source=" & OLAPServerName & _ 
      ";Initial Catalog=FoodMart;Provider=msolap;" 
   Session("ServerName") = OLAPServerName 
   Session("InitialCatalog") = "FoodMart" 
End if 
 
If Len(Session("CubeName")) > 0 Then 
   Set cdf = cat.CubeDefs(Session("CubeName")) 
Else 
   Set cdf = cat.CubeDefs("Sales") 
   Session("CubeName")="Sales" 
End if 
 
'************************************************************************ 
'*** Collect Information in HTML Form 
'************************************************************************ 
%> 
<form action="ASPADOCubeDoc.asp" method="post" id="form1" name="form1"> 
<table> 
   <tr> 
      <td> 
      <b>Olap Server name:  </b><br><input type="text" id="strServerName" name="strServerName" value="<%=Session("ServerName")%>" size="20"><br> 
 
      <b>Catalog Name:  </b><br><input type="text" id="strCatalogName" name="strCatalogName" value="<%=Session("CatalogName")%>" size="20"><br> 
 
      <b>Cube Name:  </b><br><input type="text" id="strCubeName" name="strCubeName" value="<%=Session("CubeName")%>" size="20"> 
      </td> 
      <td <TD> 
         <b>Add Property Detail:  </b><br> 
         Dimension Detail: <input type="checkbox" id="chkDimension" name="chkDimension"><br> 
 
         Hierarchy Detail: <input type="checkbox" id="chkHierarchy" name="chkHierarchy"><br> 
 
         Level Detail: <input type="checkbox" id="chkLevel" name="chkLevel"> 
      </td>  
   </tr> 
</table> 
<input type="submit" value="Cube Information" id="submit1" name="submit1"><input type="reset" value="Reset" id="reset1" name="reset1"> 
</form> 
<% 
 
'************************************************************************ 
'*** Start of Report 
'************************************************************************ 
Response.Write "<H3>Report for " & Session("CubeName") & " Cube</H3>" 
Response.Write "<OL TYPE='i'>" 
 
'************************************************************************ 
'*** Show properties of Cube 
'************************************************************************ 
            For i = 0 To cdf.Properties.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=-2>" & cdf.Properties(i).Name & ": " & cdf.Properties(i).Value & "</FONT>" 
            Next 
            Response.Write "</OL>" 
            Response.Write "<UL TYPE='SQUARE'>"    
 '************************************************************************ 
'*** Loop to display Dimension Name and Properties if Check box is  
'*** Checked 
'************************************************************************ 
      For di = 0 To cdf.Dimensions.Count - 1 
         Response.Write "<LI>" 
         Response.Write "<FONT size=4><B>Dimension: " & _ 
            cdf.Dimensions(di).Name & "</B></FONT>" 
         If Request.Form("chkDimension") = "on" Then 
            Response.Write "<OL TYPE='1'>" 
            For i = 0 To cdf.Dimensions(di).Properties.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=-2>" & _ 
                  cdf.Dimensions(di).Properties(i).Name & ": " & _ 
                  cdf.Dimensions(di).Properties(i).Value & "</FONT>" 
            Next 
            Response.Write "</OL>" 
         End If 
         Response.Write "<UL TYPE= 'Circle'>" 
'************************************************************************ 
'*** Loop to display Hierarchy Name and Properties if Check box is  
'*** Checked 
'************************************************************************ 
         For hi = 0 To cdf.Dimensions(di).Hierarchies.Count - 1 
            Response.Write "<LI>" 
            Response.Write "<FONT size=3><B>Hierarchy: " & _ 
               cdf.Dimensions(di).Hierarchies(hi).Name & "</B></FONT>" 
            If Request.Form("chkHierarchy") = "on" Then 
               Response.Write "<OL TYPE='1'>" 
               For i = 0 To _ 
                  cdf.Dimensions(di).Hierarchies(hi).Properties.Count - 1 
                  Response.Write "<LI>" 
                  Response.Write "<FONT size=-2>" & _ 
                     cdf.Dimensions(di).Hierarchies(hi).Properties(i)._ 
                     Name & ": " & _ 
                     cdf.Dimensions(di).Hierarchies(hi).Properties(i)._ 
                     Value & "</FONT>" 
               Next 
               Response.Write "</OL>" 
            End If 
            Response.Write "<UL TYPE='Disc'>" 
'************************************************************************ 
'*** Loop to display Level Name and Properties if Check box is Checked 
'************************************************************************ 
      For le = 0 To cdf.Dimensions(di).Hierarchies(hi).Levels.Count - 1 
               Response.Write "<LI>" 
               Response.Write "<FONT size=2><B>Level: " & _ 
                  cdf.Dimensions(di).Hierarchies(hi).Levels(le).Name & _ 
                  " with a Member Count of: " & _ 
                  cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                  Properties("LEVEL_CARDINALITY") & "</B></FONT>" 
               If Request.Form("chkLevel") = "on" Then 
                  Response.Write "<OL TYPE='1'>" 
                  For i = 0 To  
                     cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                     Properties.Count - 1 
                     Response.Write "<LI>" 
                     Response.Write "<FONT size=-2>" & _ 
                        cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                        Properties(i).Name & ": " & _ 
                        cdf.Dimensions(di).Hierarchies(hi).Levels(le)._ 
                        Properties(i).Value & "</FONT>" 
                  Next 
                  Response.Write "</OL>" 
               End If 
            Next 
            Response.Write "</UL>" 
         Next 
         Response.Write "</UL>" 
      Next 
      Response.Write "</UL>" 
%> 
</body> 
</html> 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]