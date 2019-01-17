---
title: Cellset example (VB)
ms.prod: access
ms.assetid: 59de67e4-0522-f52e-3e5e-2a0df500e343
ms.date: 06/08/2017
localization_priority: Normal
---


# Cellset example (VB)

**Applies to:** Access 2013 | Access 2016

This Visual Basic project demonstrates the basics of using ADO MD to access cube data. It displays member captions for column and row headers, then displays formatted values of specific cells within the cellset.


```vb
 
Sub cmdCellSettoDebugWindow_Click() 
 On Error GoTo Error_cmdCellSettoDebugWindow_Click 
 
 Dim cat As New ADOMD.Catalog 
 Dim cst As New ADOMD.CellSet 
 Dim strServer As String 
 Dim strSource As String 
 Dim strColumnHeader As String 
 Dim strRowText As String 
 Dim i As Integer 
 Dim j As Integer 
 Dim k As Integer 
 
 Screen.MousePointer = vbHourglass 
 
 '*----------------------------------------------------------------------- 
 '* Set Server to Local Host 
 '*----------------------------------------------------------------------- 
 strServer = "localhost" 
 
 '*----------------------------------------------------------------------- 
 '* Set MDX query string Source 
 '*----------------------------------------------------------------------- 
 strSource = "SELECT {[Measures].members} ON COLUMNS," & _ 
 "NON EMPTY [Store].[Store City].members ON ROWS FROM Sales" 
 
 '*----------------------------------------------------------------------- 
 '* Set Active Connection 
 '*----------------------------------------------------------------------- 
 cat.ActiveConnection = "Data Source=" & strServer & ";Provider=msolap;" 
 
 '*----------------------------------------------------------------------- 
 '* Set Cell Set source to MDX query string 
 '*----------------------------------------------------------------------- 
 cst.Source = strSource 
 
 '*----------------------------------------------------------------------- 
 '* Set Cell Sets active connection to current connection 
 '*----------------------------------------------------------------------- 
 Set cst.ActiveConnection = cat.ActiveConnection 
 
 '*----------------------------------------------------------------------- 
 '* Open Cell Set 
 '*----------------------------------------------------------------------- 
 cst.Open 
 
 '*----------------------------------------------------------------------- 
 '* Allow space for Row Header Text 
 '*----------------------------------------------------------------------- 
 strColumnHeader = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab 
 
 '*----------------------------------------------------------------------- 
 '* Loop through Column Headers 
 '*----------------------------------------------------------------------- 
 For i = 0 To cst.Axes(0).Positions.Count - 1 
 strColumnHeader = strColumnHeader & _ 
 cst.Axes(0).Positions(i).Members(0).Caption & vbTab & _ 
 vbTab & vbTab & vbTab 
 Next 
 Debug.Print vbTab & strColumnHeader & vbCrLf 
 
 '*----------------------------------------------------------------------- 
 '* Loop through Row Headers and Provide data for each row 
 '*----------------------------------------------------------------------- 
 strRowText = "" 
 For j = 0 To cst.Axes(1).Positions.Count - 1 
 strRowText = strRowText & _ 
 cst.Axes(1).Positions(j).Members(0).Caption & vbTab & _ 
 vbTab & vbTab & vbTab 
 For k = 0 To cst.Axes(0).Positions.Count - 1 
 strRowText = strRowText & cst(k, j).FormattedValue & _ 
 vbTab & vbTab & vbTab & vbTab 
 Next 
 Debug.Print strRowText & vbCrLf 
 strRowText = "" 
 Next 
 
 Screen.MousePointer = vbDefault 
 
 Exit Sub 
 
Error_cmdCellSettoDebugWindow_Click: 
 Beep 
 Screen.MousePointer = vbDefault 
 Set cat = Nothing 
 Set cst = Nothing 
 MsgBox "The Following Error has occurred:" & vbCrLf & _ 
 Err.Description, vbCritical, " Error!" 
 Exit Sub 
End Sub 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]