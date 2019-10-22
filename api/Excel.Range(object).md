---
title: Range object (Excel)
keywords: vbaxl10.chm143072
f1_keywords:
- vbaxl10.chm143072
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: b8207778-0dcc-4570-1234-f130532cc8cd
ms.date: 08/14/2019
localization_priority: Priority
---


# Range object (Excel)

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3D range.

[!include[Add-ins note](~/includes/addinsnote.md)]

## Remarks

The default member of **Range** forwards calls without parameters to the **[Value](Excel.Range.Value.md)** property and calls with parameters to the **[Item](Excel.Range.Item.md)** member. Accordingly, `someRange = someOtherRange` is equivalent to `someRange.Value = someOtherRange.Value`, `someRange(1)` to `someRange.Item(1)` and `someRange(1,1)` to `someRange.Item(1,1)`.

The following properties and methods for returning a **Range** object are described in the **Example** section:

- **[Range](Excel.Worksheet.Range.md)** and **[Cells](Excel.Worksheet.Cells.md)** properties of the **Worksheet** object
- **[Range](excel.range.range.md)** and **[Cells](excel.range.cells.md)** properties of the **Range** object   
- **[Rows](Excel.Worksheet.Rows.md)** and **[Columns](Excel.Worksheet.Columns.md)** properties of the **Worksheet** object
- **[Rows](Excel.Range.Rows.md)** and **[Columns](Excel.Range.Columns.md)** properties of the **Range** object  
- **[Offset](Excel.Range.Offset.md)** property of the **Range** object  
- **[Union](Excel.Application.Union.md)** method of the **Application** object

## Example

Use **Range** (_arg_), where _arg_ names the range, to return a **Range** object that represents a single cell or a range of cells. The following example places the value of cell A1 in cell A5.

```vb
Worksheets("Sheet1").Range("A5").Value = _ 
    Worksheets("Sheet1").Range("A1").Value
```

<br/>

The following example fills the range A1:H8 with random numbers by setting the formula for each cell in the range. When it's used without an object qualifier (an object to the left of the period), the **Range** property returns a range on the active sheet. If the active sheet isn't a worksheet, the method fails. 

Use the **[Activate](Excel.Worksheet.Activate(method).md)** method of the **Worksheet** object to activate a worksheet before you use the **Range** property without an explicit object qualifier.

```vb
Worksheets("Sheet1").Activate 
Range("A1:H8").Formula = "=Rand()"    'Range is on the active sheet
```

<br/>

The following example clears the contents of the range named _Criteria_.

> [!NOTE] 
> If you use a text argument for the range address, you must specify the address in A1-style notation (you cannot use R1C1-style notation).

```vb
Worksheets(1).Range("Criteria").ClearContents
```

<br/>

Use **Cells** on a worksheet to obtain a range consisting all single cells on the worksheet. You can access single cells via **Item**(_row_, _column_), where _row_ is the row index and _column_ is the column index. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**. 
The following example sets the value of cell A1 to 24 and of cell B1 to 42 on the first sheet of the active workbook.

```vb
Worksheets(1).Cells(1, 1).Value = 24
Worksheets(1).Cells.Item(1, 2).Value = 42
```

<br/>

The following example sets the formula for cell A2.

```vb
ActiveSheet.Cells(2, 1).Formula = "=Sum(B1:B5)"
```

<br/>

Although you can also use `Range("A1")` to return cell A1, there may be times when the **Cells** property is more convenient because you can use a variable for the row or column. The following example creates column and row headings on Sheet1. Be aware that after the worksheet has been activated, the **Cells** property can be used without an explicit sheet declaration (it returns a cell on the active sheet).

> [!NOTE] 
> Although you could use Visual Basic string functions to alter A1-style references, it is easier (and better programming practice) to use the `Cells(1, 1)` notation.

```vb
Sub SetUpTable() 
Worksheets("Sheet1").Activate 
For TheYear = 1 To 5 
    Cells(1, TheYear + 1).Value = 1990 + TheYear 
Next TheYear 
For TheQuarter = 1 To 4 
    Cells(TheQuarter + 1, 1).Value = "Q" & TheQuarter 
Next TheQuarter 
End Sub
```

<br/>

Use_expression_.**Cells**, where _expression_ is an expression that returns a **Range** object, to obtain a range with the same address consisting of single cells.
On such a range, you access single cells via **Item**(_row_, _column_), where are relative to the upper-left corner of the first area of the range. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**.
The following example sets the formula for cell C5 and D5 of the first sheet of the active workbook.

```vb
Worksheets(1).Range("C5:C10").Cells(1, 1).Formula = "=Rand()"
Worksheets(1).Range("C5:C10").Cells.Item(1, 2).Formula = "=Rand()"
```

<br/>

Use **Range** (_cell1, cell2_), where _cell1_ and _cell2_ are **Range** objects that specify the start and end cells, to return a **Range** object. The following example sets the border line style for cells A1:J10.

> [!NOTE] 
> Be aware that the period in front of each occurrence of the **Cells** property is required if the result of the preceding **With** statement is to be applied to the **Cells** property. In this case, it indicates that the cells are on worksheet one (without the period, the **Cells** property would return cells on the active sheet).

```vb
With Worksheets(1) 
    .Range(.Cells(1, 1), _ 
        .Cells(10, 10)).Borders.LineStyle = xlThick 
End With
```

<br/>

Use **Rows** on a worksheet to obtain a range consisting all rows on the worksheet. You can access single rows via **Item**(_row_), where _row_ is the row index. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**. 

> [!NOTE] 
> It is not legal to provide the second parameter of **Item** for ranges consisting of rows. You first have to convert it to single cells via **Cells**. 

The following example deletes row 4 and 10 of the first sheet of the active workbook.

```vb
Worksheets(1).Rows(10).Delete
Worksheets(1).Rows.Item(5).Delete
```

<br/>

Use **Columns** on a worksheet to obtain a range consisting all columns on the worksheet. You can access single columns via **Item**(_row_) [sic], where _row_ is the column index given as a number or as an A1-style column address. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**. 

> [!NOTE] 
> It is not legal to provide the second parameter of **Item** for ranges consisting of columns. You first have to convert it to single cells via **Cells**.

The following example deletes column "B", "C", "E", and "J" of the first sheet of the active workbook.

```vb
Worksheets(1).Columns(10).Delete
Worksheets(1).Columns.Item(5).Delete
Worksheets(1).Columns("C").Delete
Worksheets(1).Columns.Item("B").Delete
```

<br/>

Use_expression_.**Rows**, where _expression_ is an expression that returns a **Range** object, to obtain a range consisting of the rows in the first area of the range.
You can access single rows via **Item**(_row_), where _row_ is the relative row index from the top of the first area of the range. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**.

> [!NOTE] 
> It is not legal to provide the second parameter of **Item** for ranges consisting of rows. You first have to convert it to single cells via **Cells**.

The following example deletes the ranges C8:D8 and C6:D6 of the first sheet of the active workbook.

```vb
Worksheets(1).Range("C5:D10").Rows(4).Delete
Worksheets(1).Range("C5:D10").Rows.Item(2).Delete
```

<br/>

Use_expression_.**Columns**, where _expression_ is an expression that returns a **Range** object, to obtain a range consisting of the columns in the first area of the range.
You can access single columns via **Item**(_row_) [sic], where _row_ is the relative column index from the left of the first area of the range given as a number or as an A1-style column address. 
**Item** can be omitted since the call is forwarded to it by the default member of **Range**.

> [!NOTE] 
> It is not legal to provide the second parameter of **Item** for ranges consisting of columns. You first have to convert it to single cells via **Cells**.

The following example deletes the ranges L2:L10, G2:G10, F2:F10 and D2:D10 of the first sheet of the active workbook.

```vb
Worksheets(1).Range("C5:Z10").Columns(10).Delete
Worksheets(1).Range("C5:Z10").Columns.Item(5).Delete
Worksheets(1).Range("C5:Z10").Columns("D").Delete
Worksheets(1).Range("C5:Z10").Columns.Item("B").Delete
```

<br/>

Use **Offset** (_row, column_), where _row_ and _column_ are the row and column offsets, to return a range at a specified offset to another range. The following example selects the cell three rows down from and one column to the right of the cell in the upper-left corner of the current selection. You cannot select a cell that is not on the active sheet, so you must first activate the worksheet.

```vb
Worksheets("Sheet1").Activate 
  'Can't select unless the sheet is active 
Selection.Offset(3, 1).Range("A1").Select
```

<br/>

Use **Union** (_range1, range2_, ...) to return multiple-area ranges—that is, ranges composed of two or more contiguous blocks of cells. The following example creates an object defined as the union of ranges A1:B2 and C3:D4, and then selects the defined range.

```vb
Dim r1 As Range, r2 As Range, myMultiAreaRange As Range 
Worksheets("sheet1").Activate 
Set r1 = Range("A1:B2") 
Set r2 = Range("C3:D4") 
Set myMultiAreaRange = Union(r1, r2) 
myMultiAreaRange.Select
```

<br/>

If you work with selections that contain more than one area, the **[Areas](Excel.Range.Areas.md)** property is useful. It divides a multiple-area selection into individual **Range** objects and then returns the objects as a collection. You can use the **[Count](Excel.Range.Count.md)** property on the returned collection to verify a selection that contains more than one area, as shown in the following example.

```vb
Sub NoMultiAreaSelection() 
    NumberOfSelectedAreas = Selection.Areas.Count 
    If NumberOfSelectedAreas > 1 Then 
        MsgBox "You cannot carry out this command " & _ 
            "on multi-area selections" 
    End If 
End Sub
```

<br/>

This example uses the **AdvancedFilter** method of the **Range** object to create a list of the unique values, and the number of times those unique values occur, in the range of column A.

```vb
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
            Application.Evaluate("COUNTIF(" & _
            rnSource.Address(External:=True) & _
            ",""" & rnUnique(lnCount, 1).Text & """)")
    Next lnCount
    
    'Label the column of occurrences with "Occurrences"
    With rnTarget.Offset(0, 1)
        .Value = "Occurrences"
        .Font.Bold = True
    End With

End Sub
```


## Methods

- [Activate](Excel.Range.Activate.md)
- [AddComment](Excel.Range.AddComment.md)
- [AddCommentThreaded](Excel.Range.AddCommentThreaded.md)
- [AdvancedFilter](Excel.Range.AdvancedFilter.md)
- [AllocateChanges](Excel.Range.AllocateChanges.md)
- [ApplyNames](Excel.Range.ApplyNames.md)
- [ApplyOutlineStyles](Excel.Range.ApplyOutlineStyles.md)
- [AutoComplete](Excel.Range.AutoComplete.md)
- [AutoFill](Excel.Range.AutoFill.md)
- [AutoFilter](Excel.Range.AutoFilter.md)
- [AutoFit](Excel.Range.AutoFit.md)
- [AutoOutline](Excel.Range.AutoOutline.md)
- [BorderAround](Excel.Range.BorderAround.md)
- [Calculate](Excel.Range.Calculate.md)
- [CalculateRowMajorOrder](Excel.Range.CalculateRowMajorOrder.md)
- [CheckSpelling](Excel.Range.CheckSpelling.md)
- [Clear](Excel.Range.Clear.md)
- [ClearComments](Excel.Range.ClearComments.md)
- [ClearContents](Excel.Range.ClearContents.md)
- [ClearFormats](Excel.Range.ClearFormats.md)
- [ClearHyperlinks](Excel.Range.ClearHyperlinks.md)
- [ClearNotes](Excel.Range.ClearNotes.md)
- [ClearOutline](Excel.Range.ClearOutline.md)
- [ColumnDifferences](Excel.Range.ColumnDifferences.md)
- [Consolidate](Excel.Range.Consolidate.md)
- [ConvertToLinkedDataType](excel.range.converttolinkeddatatype.md)
- [Copy](Excel.Range.Copy.md)
- [CopyFromRecordset](Excel.Range.CopyFromRecordset.md)
- [CopyPicture](Excel.Range.CopyPicture.md)
- [CreateNames](Excel.Range.CreateNames.md)
- [Cut](Excel.Range.Cut.md)
- [DataTypeToText](excel.range.datatypetotext.md)
- [DataSeries](Excel.Range.DataSeries.md)
- [Delete](Excel.Range.Delete.md)
- [DialogBox](Excel.Range.DialogBox.md)
- [Dirty](Excel.Range.Dirty.md)
- [DiscardChanges](Excel.Range.DiscardChanges.md)
- [EditionOptions](Excel.Range.EditionOptions.md)
- [ExportAsFixedFormat](Excel.Range.ExportAsFixedFormat.md)
- [FillDown](Excel.Range.FillDown.md)
- [FillLeft](Excel.Range.FillLeft.md)
- [FillRight](Excel.Range.FillRight.md)
- [FillUp](Excel.Range.FillUp.md)
- [Find](Excel.Range.Find.md)
- [FindNext](Excel.Range.FindNext.md)
- [FindPrevious](Excel.Range.FindPrevious.md)
- [FlashFill](Excel.range.flashfill.md)
- [FunctionWizard](Excel.Range.FunctionWizard.md)
- [Group](Excel.Range.Group.md)
- [Insert](Excel.Range.Insert.md)
- [InsertIndent](Excel.Range.InsertIndent.md)
- [Justify](Excel.Range.Justify.md)
- [ListNames](Excel.Range.ListNames.md)
- [Merge](Excel.Range.Merge.md)
- [NavigateArrow](Excel.Range.NavigateArrow.md)
- [NoteText](Excel.Range.NoteText.md)
- [Parse](Excel.Range.Parse.md)
- [PasteSpecial](Excel.Range.PasteSpecial.md)
- [PrintOut](Excel.Range.PrintOut.md)
- [PrintPreview](Excel.Range.PrintPreview.md)
- [RemoveDuplicates](Excel.Range.RemoveDuplicates.md)
- [RemoveSubtotal](Excel.Range.RemoveSubtotal.md)
- [Replace](Excel.Range.Replace.md)
- [RowDifferences](Excel.Range.RowDifferences.md)
- [Run](Excel.Range.Run.md)
- [Select](Excel.Range.Select.md)
- [SetCellDataTypeFromCell](excel.range.setcelldatatypefromcell.md)
- [SetPhonetic](Excel.Range.SetPhonetic.md)
- [Show](Excel.Range.Show.md)
- [ShowCard](excel.range.showcard.md)
- [ShowDependents](Excel.Range.ShowDependents.md)
- [ShowErrors](Excel.Range.ShowErrors.md)
- [ShowPrecedents](Excel.Range.ShowPrecedents.md)
- [Sort](Excel.Range.Sort.md)
- [SortSpecial](Excel.Range.SortSpecial.md)
- [Speak](Excel.Range.Speak.md)
- [SpecialCells](Excel.Range.SpecialCells.md)
- [SubscribeTo](Excel.Range.SubscribeTo.md)
- [Subtotal](Excel.Range.Subtotal.md)
- [Table](Excel.Range.Table.md)
- [TextToColumns](Excel.Range.TextToColumns.md)
- [Ungroup](Excel.Range.Ungroup.md)
- [UnMerge](Excel.Range.UnMerge.md)

## Properties

- [AddIndent](Excel.Range.AddIndent.md)
- [Address](Excel.Range.Address.md)
- [AddressLocal](Excel.Range.AddressLocal.md)
- [AllowEdit](Excel.Range.AllowEdit.md)
- [Application](Excel.Range.Application.md)
- [Areas](Excel.Range.Areas.md)
- [Borders](Excel.Range.Borders.md)
- [Cells](Excel.Range.Cells.md)
- [Characters](Excel.Range.Characters.md)
- [Column](Excel.Range.Column.md)
- [Columns](Excel.Range.Columns.md)
- [ColumnWidth](Excel.Range.ColumnWidth.md)
- [Comment](Excel.Range.Comment.md)
- [CommentThreaded](Excel.Range.CommentThreaded.md)
- [Count](Excel.Range.Count.md)
- [CountLarge](Excel.Range.CountLarge.md)
- [Creator](Excel.Range.Creator.md)
- [CurrentArray](Excel.Range.CurrentArray.md)
- [CurrentRegion](Excel.Range.CurrentRegion.md)
- [Dependents](Excel.Range.Dependents.md)
- [DirectDependents](Excel.Range.DirectDependents.md)
- [DirectPrecedents](Excel.Range.DirectPrecedents.md)
- [DisplayFormat](Excel.Range.DisplayFormat.md)
- [End](Excel.Range.End.md)
- [EntireColumn](Excel.Range.EntireColumn.md)
- [EntireRow](Excel.Range.EntireRow.md)
- [Errors](Excel.Range.Errors.md)
- [Font](Excel.Range.Font.md)
- [FormatConditions](Excel.Range.FormatConditions.md)
- [Formula](Excel.Range.Formula.md)
- [FormulaArray](Excel.Range.FormulaArray.md)
- [FormulaHidden](Excel.Range.FormulaHidden.md)
- [FormulaLocal](Excel.Range.FormulaLocal.md)
- [FormulaR1C1](Excel.Range.FormulaR1C1.md)
- [FormulaR1C1Local](Excel.Range.FormulaR1C1Local.md)
- [HasArray](Excel.Range.HasArray.md)
- [HasFormula](Excel.Range.HasFormula.md)
- [HasRichDataType](excel.range.hasrichdatatype.md)
- [Height](Excel.Range.Height.md)
- [Hidden](Excel.Range.Hidden.md)
- [HorizontalAlignment](Excel.Range.HorizontalAlignment.md)
- [Hyperlinks](Excel.Range.Hyperlinks.md)
- [ID](Excel.Range.ID.md)
- [IndentLevel](Excel.Range.IndentLevel.md)
- [Interior](Excel.Range.Interior.md)
- [Item](Excel.Range.Item.md)
- [Left](Excel.Range.Left.md)
- [LinkedDataTypeState](excel.range.linkeddatatypestate.md)
- [ListHeaderRows](Excel.Range.ListHeaderRows.md)
- [ListObject](Excel.Range.ListObject.md)
- [LocationInTable](Excel.Range.LocationInTable.md)
- [Locked](Excel.Range.Locked.md)
- [MDX](Excel.Range.MDX.md)
- [MergeArea](Excel.Range.MergeArea.md)
- [MergeCells](Excel.Range.MergeCells.md)
- [Name](Excel.Range.Name.md)
- [Next](Excel.Range.Next.md)
- [NumberFormat](Excel.Range.NumberFormat.md)
- [NumberFormatLocal](Excel.Range.NumberFormatLocal.md)
- [Offset](Excel.Range.Offset.md)
- [Orientation](Excel.Range.Orientation.md)
- [OutlineLevel](Excel.Range.OutlineLevel.md)
- [PageBreak](Excel.Range.PageBreak.md)
- [Parent](Excel.Range.Parent.md)
- [Phonetic](Excel.Range.Phonetic.md)
- [Phonetics](Excel.Range.Phonetics.md)
- [PivotCell](Excel.Range.PivotCell.md)
- [PivotField](Excel.Range.PivotField.md)
- [PivotItem](Excel.Range.PivotItem.md)
- [PivotTable](Excel.Range.PivotTable.md)
- [Precedents](Excel.Range.Precedents.md)
- [PrefixCharacter](Excel.Range.PrefixCharacter.md)
- [Previous](Excel.Range.Previous.md)
- [QueryTable](Excel.Range.QueryTable.md)
- [Range](Excel.Range.Range.md)
- [ReadingOrder](Excel.Range.ReadingOrder.md)
- [Resize](Excel.Range.Resize.md)
- [Row](Excel.Range.Row.md)
- [RowHeight](Excel.Range.RowHeight.md)
- [Rows](Excel.Range.Rows.md)
- [ServerActions](Excel.Range.ServerActions.md)
- [ShowDetail](Excel.Range.ShowDetail.md)
- [ShrinkToFit](Excel.Range.ShrinkToFit.md)
- [SoundNote](Excel.Range.SoundNote.md)
- [SparklineGroups](Excel.Range.SparklineGroups.md)
- [Style](Excel.Range.Style.md)
- [Summary](Excel.Range.Summary.md)
- [Text](Excel.Range.Text.md)
- [Top](Excel.Range.Top.md)
- [UseStandardHeight](Excel.Range.UseStandardHeight.md)
- [UseStandardWidth](Excel.Range.UseStandardWidth.md)
- [Validation](Excel.Range.Validation.md)
- [Value](Excel.Range.Value.md)
- [Value2](Excel.Range.Value2.md)
- [VerticalAlignment](Excel.Range.VerticalAlignment.md)
- [Width](Excel.Range.Width.md)
- [Worksheet](Excel.Range.Worksheet.md)
- [WrapText](Excel.Range.WrapText.md)
- [XPath](Excel.Range.XPath.md)



## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
