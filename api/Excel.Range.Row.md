---
title: Range.Row property (Excel)
keywords: vbaxl10.chm144188
f1_keywords:
- vbaxl10.chm144188
ms.prod: excel
api_name:
- Excel.Range.Row
ms.assetid: 3c8d7351-4fc6-748b-c2a8-de3dab4b964e
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Row property (Excel)

Returns the number of the first row of the first area in the range. Read-only **Long**.


## Syntax

_expression_.**Row**

_expression_ A variable that represents a **[Range](Excel.Range(object).md)** object.


## Example

This example sets the row height of every other row on Sheet1 to 4 [points](../language/glossary/vbe-glossary.md#point).

```vb
For Each rw In Worksheets("Sheet1").Rows 
    If rw.Row Mod 2 = 0 Then 
        rw.RowHeight = 4 
    End If 
Next rw
```

<br/>

This example uses the **BeforeDoubleClick** worksheet event to copy a row of data from one worksheet to another. To run this code, the name of the target worksheet must be in column A. When you double-click a cell that contains data, this example gets the target worksheet name from column A and copies the entire row of data into the next available row on the target worksheet. This example accesses the active row by using the **Target** keyword.

```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'If the double click occurs on the header row or an empty cell, exit the macro.
    If Target.Row = 1 Then Exit Sub
    If Target.Row > ActiveSheet.UsedRange.Rows.Count Then Exit Sub
    If Target.Column > ActiveSheet.UsedRange.Columns.Count Then Exit Sub
    
    'Override the default double-click behavior with this function.
    Cancel = True
    
    'Declare your variables.
    Dim wks As Worksheet, xRow As Long
    
    'If an error occurs, use inline error handling.
    On Error Resume Next
    
    'Set the target worksheet as the worksheet whose name is listed in the first cell of the current row.
    Set wks = Worksheets(CStr(Cells(Target.Row, 1).Value))
    'If there is an error, exit the macro.
    If Err > 0 Then
        Err.Clear
        Exit Sub
    'Otherwise, find the next empty row in the target worksheet and copy the data into that row.
    Else
        xRow = wks.Cells(wks.Rows.Count, 1).End(xlUp).Row + 1
        wks.Range(wks.Cells(xRow, 1), wks.Cells(xRow, 7)).Value = _
        Range(Cells(Target.Row, 1), Cells(Target.Row, 7)).Value
    End If
End Sub
```


<br/>

This example deletes the empty rows from a selected range.

```vb
Sub Delete_Empty_Rows()
    'The range from which to delete the rows.
    Dim rnSelection As Range
    
    'Row and count variables used in the deletion process.
    Dim lnLastRow As Long
    Dim lnRowCount As Long
    Dim lnDeletedRows As Long
    
    'Initialize the number of deleted rows.
    lnDeletedRows = 0
    
    'Confirm that a range is selected, and that the range is contiguous.
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            
            'Initialize the range to what the user has selected, and initialize the count for the upcoming FOR loop.
            Set rnSelection = Application.Selection
            lnLastRow = rnSelection.Rows.Count
        
            'Start at the bottom row and work up: if the row is empty then
            'delete the row and increment the deleted row count.
            For lnRowCount = lnLastRow To 1 Step -1
                If Application.CountA(rnSelection.Rows(lnRowCount)) = 0 Then
                    rnSelection.Rows(lnRowCount).Delete
                    lnDeletedRows = lnDeletedRows + 1
                End If
            Next lnRowCount
        
            rnSelection.Resize(lnLastRow - lnDeletedRows).Select
         Else
            MsgBox "Please select only one area.", vbInformation
         End If
    Else
        MsgBox "Please select a range.", vbInformation
    End If
    
    'Turn screen updating back on.
    Application.ScreenUpdating = True

End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
