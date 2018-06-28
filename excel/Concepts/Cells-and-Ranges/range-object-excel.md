---
title: Range Object (Excel)
keywords: vbaxl10.chm143072
f1_keywords:
- vbaxl10.chm143072
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: b8207778-0dcc-4570-1234-f130532cc8cd
ms.date: 06/08/2017
---


# Range Object (Excel)

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.


## Example

Use  **Range** ( _arg_ ), where _arg_ names the range, to return a **Range** object that represents a single cell or a range of cells. The following example places the value of cell A1 in cell A5.


```
Worksheets("Sheet1").Range("A5").Value = _ 
    Worksheets("Sheet1").Range("A1").Value
```

The following example fills the range A1:H8 with random numbers by setting the formula for each cell in the range. When it's used without an object qualifier (an object to the left of the period), the  **Range** property returns a range on the active sheet. If the active sheet isn't a worksheet, the method fails. Use the **[Activate](../../../api/Excel.Worksheet.Activate(method).md)** method to activate a worksheet before you use the **Range** property without an explicit object qualifier.




```
Worksheets("Sheet1").Activate 
Range("A1:H8").Formula = "=Rand()"    'Range is on the active sheet
```

The following example clears the contents of the range named  _Criteria_.


 **Note**  If you use a text argument for the range address, you must specify the address in A1-style notation (you cannot use R1C1-style notation).




```
Worksheets(1).Range("Criteria").ClearContents
```

Use  **Cells** ( _row_, _column_ ) where _row_ is the row index and _column_ is the column index, to return a single cell. The following example sets the value of cell A1 to 24.




```
Worksheets(1).Cells(1, 1).Value = 24
```

The following example sets the formula for cell A2.




```
ActiveSheet.Cells(2, 1).Formula = "=Sum(B1:B5)"
```

Although you can also use  `Range("A1")` to return cell A1, there may be times when the **Cells** property is more convenient because you can use a variable for the row or column. The following example creates column and row headings on Sheet1. Be aware that after the worksheet has been activated, the **Cells** property can be used without an explicit sheet declaration (it returns a cell on the active sheet).


 **Note**  Although you could use Visual Basic string functions to alter A1-style references, it is easier (and better programming practice) to use the  `Cells(1, 1)` notation.




```
Sub SetUpTable() 
Worksheets("Sheet1").Activate 
For TheYear = 1 To 5 
    Cells(1, TheYear + 1).Value = 1990 + TheYear 
Next TheYear 
For TheQuarter = 1 To 4 
    Cells(TheQuarter + 1, 1).Value = "Q" &amp; TheQuarter 
Next TheQuarter 
End Sub
```

Use  _expression_. **Cells** ( _row_, _column_ ), where _expression_ is an expression that returns a **Range** object, and _row_ and _column_ are relative to the upper-left corner of the range, to return part of a range. The following example sets the formula for cell C5.




```
Worksheets(1).Range("C5:C10").Cells(1, 1).Formula = "=Rand()"
```

Use  **Range** ( _cell1, cell2_ ), where _cell1_ and _cell2_ are **Range** objects that specify the start and end cells, to return a **Range** object. The following example sets the border line style for cells A1:J10.


 **Note**  Be aware that the period in front of each occurrence of the  **Cells** property. The period is required if the result of the preceding **With** statement is to be applied to the **Cells** property—in this case, to indicate that the cells are on worksheet one (without the period, the **Cells** property would return cells on the active sheet).




```
With Worksheets(1) 
    .Range(.Cells(1, 1), _ 
        .Cells(10, 10)).Borders.LineStyle = xlThick 
End With
```

Use  **Offset** ( _row, column_ ), where _row_ and _column_ are the row and column offsets, to return a range at a specified offset to another range. The following example selects the cell three rows down from and one column to the right of the cell in the upper-left corner of the current selection. You cannot select a cell that is not on the active sheet, so you must first activate the worksheet.




```
Worksheets("Sheet1").Activate 
  'Can't select unless the sheet is active 
Selection.Offset(3, 1).Range("A1").Select
```

Use  **Union** ( _range1, range2_, ...) to return multiple-area ranges—that is, ranges composed of two or more contiguous blocks of cells. The following example creates an object defined as the union of ranges A1:B2 and C3:D4, and then selects the defined range.




```
Dim r1 As Range, r2 As Range, myMultiAreaRange As Range 
Worksheets("sheet1").Activate 
Set r1 = Range("A1:B2") 
Set r2 = Range("C3:D4") 
Set myMultiAreaRange = Union(r1, r2) 
myMultiAreaRange.Select
```

If you work with selections that contain more than one area, the  **[Areas](../../../api/Excel.Range.Areas.md)** property is useful. It divides a multiple-area selection into individual **Range** objects and then returns the objects as a collection. You can use the **[Count](../../../api/Excel.Range.Count.md)** property on the returned collection to verify a selection that contains more than one area, as shown in the following example.




```
Sub NoMultiAreaSelection() 
    NumberOfSelectedAreas = Selection.Areas.Count 
    If NumberOfSelectedAreas > 1 Then 
        MsgBox "You cannot carry out this command " &amp; _ 
            "on multi-area selections" 
    End If 
End Sub
```

 **Sample code provided by:** Dennis Wallentin,[VSTO &amp; .NET &amp; Excel](http://xldennis.wordpress.com/)

This example uses the  **AdvancedFilter** method of the **Range** object to create a list of the unique values, and the number of times those unique values occur, in the range of column A.




```
Sub Create_Unique_List_Count()
    'Excel workbook, the source and target worksheets, and the source and target ranges.
    Dim wbBook As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rnSource As Range
    Dim rnTarget As Range
    Dim rnUnique As Range
    'Variant to hold the unique data
    Dim vaUnique As Variant
    'Number of unique values in the data
    Dim lnCount As Long
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    With wbBook
        Set wsSource = .Worksheets("Sheet1")
        Set wsTarget = .Worksheets("Sheet2")
    End With
    
    'On the source worksheet, set the range to the data stored in column A
    With wsSource
        Set rnSource = .Range(.Range("A1"), .Range("A100").End(xlDown))
    End With
    
    'On the target worksheet, set the range as column A.
    Set rnTarget = wsTarget.Range("A1")
    
    'Use AdvancedFilter to copy the data from the source to the target,
    'while filtering for duplicate values.
    rnSource.AdvancedFilter Action:=xlFilterCopy, _
                            CopyToRange:=rnTarget, _
                            Unique:=True
                            
    'On the target worksheet, set the unique range on Column A, excluding the first cell
    '(which will contain the "List" header for the column).
    With wsTarget
        Set rnUnique = .Range(.Range("A2"), .Range("A100").End(xlUp))
    End With
    
    'Assign all the values of the Unique range into the Unique variant.
    vaUnique = rnUnique.Value
    
    'Count the number of occurrences of every unique value in the source data,
    'and list it next to its relevant value.
    For lnCount = 1 To UBound(vaUnique)
        rnUnique(lnCount, 1).Offset(0, 1).Value = _
            Application.Evaluate("COUNTIF(" &amp; _
            rnSource.Address(External:=True) &amp; _
            ",""" &amp; rnUnique(lnCount, 1).Text &amp; """)")
    Next lnCount
    
    'Label the column of occurrences with "Occurrences"
    With rnTarget.Offset(0, 1)
        .Value = "Occurrences"
        .Font.Bold = True
    End With

End Sub
```


## Remarks

The following properties and methods for returning a  **Range** object are described in the examples section:


-  **[Range](../../../api/Excel.Worksheet.Range.md)** property
    
-  **[Cells](../../../api/Excel.Worksheet.Cells.md)** property
    
-  **Range** and **Cells**
    
-  **[Offset](../../../api/Excel.Range.Offset.md)** property
    
-  **[Union](../../../api/Excel.Application.Union.md)** method
    

## Methods



|**Name**|
|:-----|
|[Activate](../../../api/Excel.Range.Activate.md)|
|[AddComment](../../../api/Excel.Range.AddComment.md)|
|[AdvancedFilter](../../../api/Excel.Range.AdvancedFilter.md)|
|[AllocateChanges](../../../api/Excel.Range.AllocateChanges.md)|
|[ApplyNames](../../../api/Excel.Range.ApplyNames.md)|
|[ApplyOutlineStyles](../../../api/Excel.Range.ApplyOutlineStyles.md)|
|[AutoComplete](../../../api/Excel.Range.AutoComplete.md)|
|[AutoFill](../../../api/Excel.Range.AutoFill.md)|
|[AutoFilter](../../../api/Excel.Range.AutoFilter.md)|
|[AutoFit](../../../api/Excel.Range.AutoFit.md)|
|[AutoOutline](../../../api/Excel.Range.AutoOutline.md)|
|[BorderAround](../../../api/Excel.Range.BorderAround.md)|
|[Calculate](../../../api/Excel.Range.Calculate.md)|
|[CalculateRowMajorOrder](../../../api/Excel.Range.CalculateRowMajorOrder.md)|
|[CheckSpelling](../../../api/Excel.Range.CheckSpelling.md)|
|[Clear](../../../api/Excel.Range.Clear.md)|
|[ClearComments](../../../api/Excel.Range.ClearComments.md)|
|[ClearContents](../../../api/Excel.Range.ClearContents.md)|
|[ClearFormats](../../../api/Excel.Range.ClearFormats.md)|
|[ClearHyperlinks](../../../api/Excel.Range.ClearHyperlinks.md)|
|[ClearNotes](../../../api/Excel.Range.ClearNotes.md)|
|[ClearOutline](../../../api/Excel.Range.ClearOutline.md)|
|[ColumnDifferences](../../../api/Excel.Range.ColumnDifferences.md)|
|[Consolidate](../../../api/Excel.Range.Consolidate.md)|
|[Copy](../../../api/Excel.Range.Copy.md)|
|[CopyFromRecordset](../../../api/Excel.Range.CopyFromRecordset.md)|
|[CopyPicture](../../../api/Excel.Range.CopyPicture.md)|
|[CreateNames](../../../api/Excel.Range.CreateNames.md)|
|[Cut](../../../api/Excel.Range.Cut.md)|
|[DataSeries](../../../api/Excel.Range.DataSeries.md)|
|[Delete](../../../api/Excel.Range.Delete.md)|
|[DialogBox](../../../api/Excel.Range.DialogBox.md)|
|[Dirty](../../../api/Excel.Range.Dirty.md)|
|[DiscardChanges](../../../api/Excel.Range.DiscardChanges.md)|
|[EditionOptions](../../../api/Excel.Range.EditionOptions.md)|
|[ExportAsFixedFormat](../../../api/Excel.Range.ExportAsFixedFormat.md)|
|[FillDown](../../../api/Excel.Range.FillDown.md)|
|[FillLeft](../../../api/Excel.Range.FillLeft.md)|
|[FillRight](../../../api/Excel.Range.FillRight.md)|
|[FillUp](../../../api/Excel.Range.FillUp.md)|
|[Find](../../../api/Excel.Range.Find.md)|
|[FindNext](../../../api/Excel.Range.FindNext.md)|
|[FindPrevious](../../../api/Excel.Range.FindPrevious.md)|
|[FlashFill](../../../api/Excel.range.flashfill.md)|
|[FunctionWizard](../../../api/Excel.Range.FunctionWizard.md)|
|[Group](../../../api/Excel.Range.Group.md)|
|[Insert](../../../api/Excel.Range.Insert.md)|
|[InsertIndent](../../../api/Excel.Range.InsertIndent.md)|
|[Justify](../../../api/Excel.Range.Justify.md)|
|[ListNames](../../../api/Excel.Range.ListNames.md)|
|[Merge](../../../api/Excel.Range.Merge.md)|
|[NavigateArrow](../../../api/Excel.Range.NavigateArrow.md)|
|[NoteText](../../../api/Excel.Range.NoteText.md)|
|[Parse](../../../api/Excel.Range.Parse.md)|
|[PasteSpecial](../../../api/Excel.Range.PasteSpecial.md)|
|[PrintOut](../../../api/Excel.Range.PrintOut.md)|
|[PrintPreview](../../../api/Excel.Range.PrintPreview.md)|
|[RemoveDuplicates](../../../api/Excel.Range.RemoveDuplicates.md)|
|[RemoveSubtotal](../../../api/Excel.Range.RemoveSubtotal.md)|
|[Replace](../../../api/Excel.Range.Replace.md)|
|[RowDifferences](../../../api/Excel.Range.RowDifferences.md)|
|[Run](../../../api/Excel.Range.Run.md)|
|[Select](../../../api/Excel.Range.Select.md)|
|[SetPhonetic](../../../api/Excel.Range.SetPhonetic.md)|
|[Show](../../../api/Excel.Range.Show.md)|
|[ShowDependents](../../../api/Excel.Range.ShowDependents.md)|
|[ShowErrors](../../../api/Excel.Range.ShowErrors.md)|
|[ShowPrecedents](../../../api/Excel.Range.ShowPrecedents.md)|
|[Sort](../../../api/Excel.Range.Sort.md)|
|[SortSpecial](../../../api/Excel.Range.SortSpecial.md)|
|[Speak](../../../api/Excel.Range.Speak.md)|
|[SpecialCells](../../../api/Excel.Range.SpecialCells.md)|
|[SubscribeTo](../../../api/Excel.Range.SubscribeTo.md)|
|[Subtotal](../../../api/Excel.Range.Subtotal.md)|
|[Table](../../../api/Excel.Range.Table.md)|
|[TextToColumns](../../../api/Excel.Range.TextToColumns.md)|
|[Ungroup](../../../api/Excel.Range.Ungroup.md)|
|[UnMerge](../../../api/Excel.Range.UnMerge.md)|

## Properties



|**Name**|
|:-----|
|[AddIndent](../../../api/Excel.Range.AddIndent.md)|
|[Address](../../../api/Excel.Range.Address.md)|
|[AddressLocal](../../../api/Excel.Range.AddressLocal.md)|
|[AllowEdit](../../../api/Excel.Range.AllowEdit.md)|
|[Application](../../../api/Excel.Range.Application.md)|
|[Areas](../../../api/Excel.Range.Areas.md)|
|[Borders](../../../api/Excel.Range.Borders.md)|
|[Cells](../../../api/Excel.Range.Cells.md)|
|[Characters](../../../api/Excel.Range.Characters.md)|
|[Column](../../../api/Excel.Range.Column.md)|
|[Columns](../../../api/Excel.Range.Columns.md)|
|[ColumnWidth](../../../api/Excel.Range.ColumnWidth.md)|
|[Comment](../../../api/Excel.Range.Comment.md)|
|[Count](../../../api/Excel.Range.Count.md)|
|[CountLarge](../../../api/Excel.Range.CountLarge.md)|
|[Creator](../../../api/Excel.Range.Creator.md)|
|[CurrentArray](../../../api/Excel.Range.CurrentArray.md)|
|[CurrentRegion](../../../api/Excel.Range.CurrentRegion.md)|
|[Dependents](../../../api/Excel.Range.Dependents.md)|
|[DirectDependents](../../../api/Excel.Range.DirectDependents.md)|
|[DirectPrecedents](../../../api/Excel.Range.DirectPrecedents.md)|
|[DisplayFormat](../../../api/Excel.Range.DisplayFormat.md)|
|[End](../../../api/Excel.Range.End.md)|
|[EntireColumn](../../../api/Excel.Range.EntireColumn.md)|
|[EntireRow](../../../api/Excel.Range.EntireRow.md)|
|[Errors](../../../api/Excel.Range.Errors.md)|
|[Font](../../../api/Excel.Range.Font.md)|
|[FormatConditions](../../../api/Excel.Range.FormatConditions.md)|
|[Formula](../../../api/Excel.Range.Formula.md)|
|[FormulaArray](../../../api/Excel.Range.FormulaArray.md)|
|[FormulaHidden](../../../api/Excel.Range.FormulaHidden.md)|
|[FormulaLocal](../../../api/Excel.Range.FormulaLocal.md)|
|[FormulaR1C1](../../../api/Excel.Range.FormulaR1C1.md)|
|[FormulaR1C1Local](../../../api/Excel.Range.FormulaR1C1Local.md)|
|[HasArray](../../../api/Excel.Range.HasArray.md)|
|[HasFormula](../../../api/Excel.Range.HasFormula.md)|
|[Height](../../../api/Excel.Range.Height.md)|
|[Hidden](../../../api/Excel.Range.Hidden.md)|
|[HorizontalAlignment](../../../api/Excel.Range.HorizontalAlignment.md)|
|[Hyperlinks](../../../api/Excel.Range.Hyperlinks.md)|
|[ID](../../../api/Excel.Range.ID.md)|
|[IndentLevel](../../../api/Excel.Range.IndentLevel.md)|
|[Interior](../../../api/Excel.Range.Interior.md)|
|[Item](../../../api/Excel.Range.Item.md)|
|[Left](../../../api/Excel.Range.Left.md)|
|[ListHeaderRows](../../../api/Excel.Range.ListHeaderRows.md)|
|[ListObject](../../../api/Excel.Range.ListObject.md)|
|[LocationInTable](../../../api/Excel.Range.LocationInTable.md)|
|[Locked](../../../api/Excel.Range.Locked.md)|
|[MDX](../../../api/Excel.Range.MDX.md)|
|[MergeArea](../../../api/Excel.Range.MergeArea.md)|
|[MergeCells](../../../api/Excel.Range.MergeCells.md)|
|[Name](../../../api/Excel.Range.Name.md)|
|[Next](../../../api/Excel.Range.Next.md)|
|[NumberFormat](../../../api/Excel.Range.NumberFormat.md)|
|[NumberFormatLocal](../../../api/Excel.Range.NumberFormatLocal.md)|
|[Offset](../../../api/Excel.Range.Offset.md)|
|[Orientation](../../../api/Excel.Range.Orientation.md)|
|[OutlineLevel](../../../api/Excel.Range.OutlineLevel.md)|
|[PageBreak](../../../api/Excel.Range.PageBreak.md)|
|[Parent](../../../api/Excel.Range.Parent.md)|
|[Phonetic](../../../api/Excel.Range.Phonetic.md)|
|[Phonetics](../../../api/Excel.Range.Phonetics.md)|
|[PivotCell](../../../api/Excel.Range.PivotCell.md)|
|[PivotField](../../../api/Excel.Range.PivotField.md)|
|[PivotItem](../../../api/Excel.Range.PivotItem.md)|
|[PivotTable](../../../api/Excel.Range.PivotTable.md)|
|[Precedents](../../../api/Excel.Range.Precedents.md)|
|[PrefixCharacter](../../../api/Excel.Range.PrefixCharacter.md)|
|[Previous](../../../api/Excel.Range.Previous.md)|
|[QueryTable](../../../api/Excel.Range.QueryTable.md)|
|[Range](../../../api/Excel.Range.Range.md)|
|[ReadingOrder](../../../api/Excel.Range.ReadingOrder.md)|
|[Resize](../../../api/Excel.Range.Resize.md)|
|[Row](../../../api/Excel.Range.Row.md)|
|[RowHeight](../../../api/Excel.Range.RowHeight.md)|
|[Rows](../../../api/Excel.Range.Rows.md)|
|[ServerActions](../../../api/Excel.Range.ServerActions.md)|
|[ShowDetail](../../../api/Excel.Range.ShowDetail.md)|
|[ShrinkToFit](../../../api/Excel.Range.ShrinkToFit.md)|
|[SoundNote](../../../api/Excel.Range.SoundNote.md)|
|[SparklineGroups](../../../api/Excel.Range.SparklineGroups.md)|
|[Style](../../../api/Excel.Range.Style.md)|
|[Summary](../../../api/Excel.Range.Summary.md)|
|[Text](../../../api/Excel.Range.Text.md)|
|[Top](../../../api/Excel.Range.Top.md)|
|[UseStandardHeight](../../../api/Excel.Range.UseStandardHeight.md)|
|[UseStandardWidth](../../../api/Excel.Range.UseStandardWidth.md)|
|[Validation](../../../api/Excel.Range.Validation.md)|
|[Value](../../../api/Excel.Range.Value.md)|
|[Value2](../../../api/Excel.Range.Value2.md)|
|[VerticalAlignment](../../../api/Excel.Range.VerticalAlignment.md)|
|[Width](../../../api/Excel.Range.Width.md)|
|[Worksheet](../../../api/Excel.Range.Worksheet.md)|
|[WrapText](../../../api/Excel.Range.WrapText.md)|
|[XPath](../../../api/Excel.Range.XPath.md)|

## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &amp; .NET &amp; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


