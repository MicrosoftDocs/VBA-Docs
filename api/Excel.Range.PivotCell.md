---
title: Range.PivotCell property (Excel)
keywords: vbaxl10.chm144233
f1_keywords:
- vbaxl10.chm144233
ms.prod: excel
api_name:
- Excel.Range.PivotCell
ms.assetid: 976f6393-db3b-d52a-0cbc-88a73bb7c070
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.PivotCell property (Excel)

Returns a  **[PivotCell](Excel.PivotCell.md)** object that represents a cell in a PivotTable report.


## Syntax

_expression_. `PivotCell`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Example

This example determines the name of the PivotTable the  **PivotCell** object is located in and notifies the user. The example assumes that a PivotTable exists on the active worksheet and that cell A3 is located in the PivotTable.


```vb
Sub CheckPivotCell() 
 
 'Determine the name of the PivotTable the PivotCell is located in. 
 MsgBox "Cell A3 is located in PivotTable: " & _ 
 Application.Range("A3").PivotCell.Parent 
 
End Sub
```


## See also


[Range Object](Excel.Range(object).md)
[ValueChange Object](Excel.ValueChange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]