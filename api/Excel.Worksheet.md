---
title: Worksheet object (Excel)
keywords: vbaxl10.chm173072
f1_keywords:
- vbaxl10.chm173072
ms.prod: excel
api_name:
- Excel.Worksheet
ms.assetid: 182b705e-854a-81cc-a4b0-59b942de55ae
ms.date: 05/15/2019
localization_priority: Normal
---


# Worksheet object (Excel)

Represents a worksheet.


## Remarks

The **Worksheet** object is a member of the **[Worksheets](Excel.Worksheets.md)** collection. The **Worksheets** collection contains all the **Worksheet** objects in a workbook.

The **Worksheet** object is also a member of the **[Sheets](Excel.Sheets.md)** collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).


## Example

Use **[Worksheets](Excel.Workbook.Worksheets.md)** (_index_), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.

```vb
Worksheets(1).Visible = False
```

The worksheet index number denotes the position of the worksheet on the workbook's tab bar. `Worksheets(1)` is the first (leftmost) worksheet in the workbook, and `Worksheets(Worksheets.Count)` is the last one. All worksheets are included in the index count, even if they are hidden.

<br/>

The worksheet name is shown on the tab for the worksheet. Use the **[Name](Excel.Worksheet.Name.md)** property to set or return the worksheet name. The following example protects the scenarios on Sheet1.

```vb
 
Dim strPassword As String 
strPassword = InputBox ("Enter the password for the worksheet") 
Worksheets("Sheet1").Protect password:=strPassword, scenarios:=True
```

<br/>

When a worksheet is the active sheet, you can use the **[ActiveSheet](Excel.Workbook.ActiveSheet.md)** property to refer to it. The following example uses the **[Activate](Excel.Worksheet.Activate(method).md)** method to activate Sheet1, sets the page orientation to landscape mode, and then prints the worksheet.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.Orientation = xlLandscape 
ActiveSheet.PrintOut
```

<br/>

This example uses the **[BeforeDoubleClick](Excel.Worksheet.BeforeDoubleClick.md)** event to open a specified set of files in Notepad. To use this example, your worksheet must contain the following data:

- Cell A1 must contain the names of the files to open, each separated by a comma and a space.    
- Cell D1 must contain the path to where the Notepad files are located.    
- Cell D2 must contain the path to where the Notepad program is located.   
- Cell D3 must contain the file extension, without the period, for the Notepad files (txt).
    
When you double-click cell A1, the files specified in cell A1 are opened in Notepad.

```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
   'Define your variables.
   Dim sFile As String, sPath As String, sTxt As String, sExe As String, sSfx As String
   
   'If you did not double-click on A1, then exit the function.
   If Target.Address <> "$A$1" Then Exit Sub
   
   'If you did double-click on A1, then override the default double-click behavior with this function.
   Cancel = True
   
   'Set the path to the files, the path to Notepad, the file extension of the files, and the names of the files,
   'based on the information on the worksheet.
   sPath = Range("D1").Value
   sExe = Range("D2").Value
   sSfx = Range("D3").Value
   sFile = Range("A1").Value
   
   'Remove the spaces between the file names.
   sFile = WorksheetFunction.Substitute(sFile, " ", "")
   
   'Go through each file in the list (separated by commas) and
   'create the path, call the executable, and move on to the next comma.
   Do While InStr(sFile, ",")
      sTxt = sPath & "\" & Left(sFile, InStr(sFile, ",") - 1) & "." & sSfx
      If Dir(sTxt) <> "" Then Shell sExe & " " & sTxt, vbNormalFocus
      sFile = Right(sFile, Len(sFile) - InStr(sFile, ","))
   Loop
   
   'Finish off the last file name in the list
   sTxt = sPath & "\" & sFile & "." & sSfx
   If Dir(sTxt) <> "" Then Shell sExe & " " & sTxt, vbNormalNoFocus
End Sub
```


## Events

- [Activate](Excel.Worksheet.Activate(even).md)
- [BeforeDelete](Excel.worksheet.beforedelete.md)
- [BeforeDoubleClick](Excel.Worksheet.BeforeDoubleClick.md)
- [BeforeRightClick](Excel.Worksheet.BeforeRightClick.md)
- [Calculate](Excel.Worksheet.Calculate(even).md)
- [Change](Excel.Worksheet.Change.md)
- [Deactivate](Excel.Worksheet.Deactivate.md)
- [FollowHyperlink](Excel.Worksheet.FollowHyperlink.md)
- [LensGalleryRenderComplete](Excel.worksheet.lensgalleryrendercomplete.md)
- [PivotTableAfterValueChange](Excel.Worksheet.PivotTableAfterValueChange.md)
- [PivotTableBeforeAllocateChanges](Excel.Worksheet.PivotTableBeforeAllocateChanges.md)
- [PivotTableBeforeCommitChanges](Excel.Worksheet.PivotTableBeforeCommitChanges.md)
- [PivotTableBeforeDiscardChanges](Excel.Worksheet.PivotTableBeforeDiscardChanges.md)
- [PivotTableChangeSync](Excel.Worksheet.PivotTableChangeSync.md)
- [PivotTableUpdate](Excel.Worksheet.PivotTableUpdate.md)
- [SelectionChange](Excel.Worksheet.SelectionChange.md)
- [TableUpdate](Excel.worksheet.tableupdate.md)

## Methods

- [Activate](Excel.Worksheet.Activate(method).md)
- [Calculate](Excel.Worksheet.Calculate(method).md)
- [ChartObjects](Excel.Worksheet.ChartObjects.md)
- [CheckSpelling](Excel.Worksheet.CheckSpelling.md)
- [CircleInvalid](Excel.Worksheet.CircleInvalid.md)
- [ClearArrows](Excel.Worksheet.ClearArrows.md)
- [ClearCircles](Excel.Worksheet.ClearCircles.md)
- [Copy](Excel.Worksheet.Copy.md)
- [Delete](Excel.Worksheet.Delete.md)
- [Evaluate](Excel.Worksheet.Evaluate.md)
- [ExportAsFixedFormat](Excel.Worksheet.ExportAsFixedFormat.md)
- [Move](Excel.Worksheet.Move.md)
- [OLEObjects](Excel.Worksheet.OLEObjects.md)
- [Paste](Excel.Worksheet.Paste.md)
- [PasteSpecial](Excel.Worksheet.PasteSpecial.md)
- [PivotTables](Excel.Worksheet.PivotTables.md)
- [PivotTableWizard](Excel.Worksheet.PivotTableWizard.md)
- [PrintOut](Excel.Worksheet.PrintOut.md)
- [PrintPreview](Excel.Worksheet.PrintPreview.md)
- [Protect](Excel.Worksheet.Protect.md)
- [ResetAllPageBreaks](Excel.Worksheet.ResetAllPageBreaks.md)
- [SaveAs](Excel.Worksheet.SaveAs.md)
- [Scenarios](Excel.Worksheet.Scenarios.md)
- [Select](Excel.Worksheet.Select.md)
- [SetBackgroundPicture](Excel.Worksheet.SetBackgroundPicture.md)
- [ShowAllData](Excel.Worksheet.ShowAllData.md)
- [ShowDataForm](Excel.Worksheet.ShowDataForm.md)
- [Unprotect](Excel.Worksheet.Unprotect.md)
- [XmlDataQuery](Excel.Worksheet.XmlDataQuery.md)
- [XmlMapQuery](Excel.Worksheet.XmlMapQuery.md)

## Properties

- [Application](Excel.Worksheet.Application.md)
- [AutoFilter](Excel.Worksheet.AutoFilter.md)
- [AutoFilterMode](Excel.Worksheet.AutoFilterMode.md)
- [Cells](Excel.Worksheet.Cells.md)
- [CircularReference](Excel.Worksheet.CircularReference.md)
- [CodeName](Excel.Worksheet.CodeName.md)
- [Columns](Excel.Worksheet.Columns.md)
- [Comments](Excel.Worksheet.Comments.md)
- [CommentsThreaded](Excel.Worksheet.CommentsThreaded.md)
- [ConsolidationFunction](Excel.Worksheet.ConsolidationFunction.md)
- [ConsolidationOptions](Excel.Worksheet.ConsolidationOptions.md)
- [ConsolidationSources](Excel.Worksheet.ConsolidationSources.md)
- [Creator](Excel.Worksheet.Creator.md)
- [CustomProperties](Excel.Worksheet.CustomProperties.md)
- [DisplayPageBreaks](Excel.Worksheet.DisplayPageBreaks.md)
- [DisplayRightToLeft](Excel.Worksheet.DisplayRightToLeft.md)
- [EnableAutoFilter](Excel.Worksheet.EnableAutoFilter.md)
- [EnableCalculation](Excel.Worksheet.EnableCalculation.md)
- [EnableFormatConditionsCalculation](Excel.Worksheet.EnableFormatConditionsCalculation.md)
- [EnableOutlining](Excel.Worksheet.EnableOutlining.md)
- [EnablePivotTable](Excel.Worksheet.EnablePivotTable.md)
- [EnableSelection](Excel.Worksheet.EnableSelection.md)
- [FilterMode](Excel.Worksheet.FilterMode.md)
- [HPageBreaks](Excel.Worksheet.HPageBreaks.md)
- [Hyperlinks](Excel.Worksheet.Hyperlinks.md)
- [Index](Excel.Worksheet.Index.md)
- [ListObjects](Excel.Worksheet.ListObjects.md)
- [MailEnvelope](Excel.Worksheet.MailEnvelope.md)
- [Name](Excel.Worksheet.Name.md)
- [Names](Excel.Worksheet.Names.md)
- [Next](Excel.Worksheet.Next.md)
- [Outline](Excel.Worksheet.Outline.md)
- [PageSetup](Excel.Worksheet.PageSetup.md)
- [Parent](Excel.Worksheet.Parent.md)
- [Previous](Excel.Worksheet.Previous.md)
- [PrintedCommentPages](Excel.Worksheet.PrintedCommentPages.md)
- [ProtectContents](Excel.Worksheet.ProtectContents.md)
- [ProtectDrawingObjects](Excel.Worksheet.ProtectDrawingObjects.md)
- [Protection](Excel.Worksheet.Protection.md)
- [ProtectionMode](Excel.Worksheet.ProtectionMode.md)
- [ProtectScenarios](Excel.Worksheet.ProtectScenarios.md)
- [QueryTables](Excel.Worksheet.QueryTables.md)
- [Range](Excel.Worksheet.Range.md)
- [Rows](Excel.Worksheet.Rows.md)
- [ScrollArea](Excel.Worksheet.ScrollArea.md)
- [Shapes](Excel.Worksheet.Shapes.md)
- [Sort](Excel.Worksheet.Sort.md)
- [StandardHeight](Excel.Worksheet.StandardHeight.md)
- [StandardWidth](Excel.Worksheet.StandardWidth.md)
- [Tab](Excel.Worksheet.Tab.md)
- [TransitionExpEval](Excel.Worksheet.TransitionExpEval.md)
- [TransitionFormEntry](Excel.Worksheet.TransitionFormEntry.md)
- [Type](Excel.Worksheet.Type.md)
- [UsedRange](Excel.Worksheet.UsedRange.md)
- [Visible](Excel.Worksheet.Visible.md)
- [VPageBreaks](Excel.Worksheet.VPageBreaks.md)



## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
