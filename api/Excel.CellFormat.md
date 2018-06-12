---
title: CellFormat Object (Excel)
keywords: vbaxl10.chm675072
f1_keywords:
- vbaxl10.chm675072
ms.prod: excel
api_name:
- Excel.CellFormat
ms.assetid: da4e50b9-6d5b-22e1-3113-0d1ea6686272
ms.date: 06/08/2017
---


# CellFormat Object (Excel)

Represents the search criteria for the cell format.


## Remarks

Use the  **[FindFormat](Excel.Application.FindFormat.md)** or **[ReplaceFormat](Excel.Application.ReplaceFormat.md)** properties of the **[Application](Excel.Application(objec).md)** object to return a **CellFormat** object.

With a  **CellFormat** object, you can use the **[Borders](Excel.CellFormat.Borders.md)**, **[Font](Excel.CellFormat.Font.md)**, or **[Interior](Excel.CellFormat.Interior.md)** properties of the **CellFormat** object, to define the search criteria for the cell format.


## Example

The following example sets the search criteria for the interior of the cell format. 


```
Sub ChangeCellFormat() 
 
 ' Set the interior of cell A1 to yellow. 
 Range("A1").Select 
 Selection.Interior.ColorIndex = 36 
 MsgBox "The cell format for cell A1 is a yellow interior." 
 
 ' Set the CellFormat object to replace yellow with green. 
 With Application 
 .FindFormat.Interior.ColorIndex = 36 
 .ReplaceFormat.Interior.ColorIndex = 35 
 End With 
 
 ' Find and replace cell A1's yellow interior with green. 
 ActiveCell.Replace What:="", Replacement:="", LookAt:=xlPart, _ 
 SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _ 
 ReplaceFormat:=True 
 MsgBox "The cell format for cell A1 is replaced with a green interior." 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Clear](Excel.CellFormat.Clear.md)|

## Properties



|**Name**|
|:-----|
|[AddIndent](Excel.CellFormat.AddIndent.md)|
|[Application](Excel.CellFormat.Application.md)|
|[Borders](Excel.CellFormat.Borders.md)|
|[Creator](Excel.CellFormat.Creator.md)|
|[Font](Excel.CellFormat.Font.md)|
|[FormulaHidden](Excel.CellFormat.FormulaHidden.md)|
|[HorizontalAlignment](Excel.CellFormat.HorizontalAlignment.md)|
|[IndentLevel](Excel.CellFormat.IndentLevel.md)|
|[Interior](Excel.CellFormat.Interior.md)|
|[Locked](Excel.CellFormat.Locked.md)|
|[MergeCells](Excel.CellFormat.MergeCells.md)|
|[NumberFormat](Excel.CellFormat.NumberFormat.md)|
|[NumberFormatLocal](Excel.CellFormat.NumberFormatLocal.md)|
|[Orientation](Excel.CellFormat.Orientation.md)|
|[Parent](Excel.CellFormat.Parent.md)|
|[ShrinkToFit](cellformat-shrinktofit-property-excel.md)|
|[VerticalAlignment](cellformat-verticalalignment-property-excel.md)|
|[WrapText](cellformat-wraptext-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
