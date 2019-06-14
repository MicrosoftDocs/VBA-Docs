---
title: Table.Cells property (Publisher)
keywords: vbapb10.chm4784136
f1_keywords:
- vbapb10.chm4784136
ms.prod: publisher
api_name:
- Publisher.Table.Cells
ms.assetid: 42622697-aef1-0765-7d85-4919c298d92f
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.Cells property (Publisher)

Returns a **[CellRange](publisher.cellrange.md)** object that represents a range of cells in a table.


## Syntax

_expression_.**Cells** (_StartRow_, _StartColumn_, _EndRow_, _EndColumn_)

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_StartRow_|Optional| **Long**|The row in which the starting cell exists. If this argument is omitted, all the table rows are included in the range.|
|_StartColumn_|Optional| **Long**|The column in which the starting cell exists. If this argument is omitted, all the table columns are included in the range.|
|_EndRow_|Optional| **Long**|The row in which the ending cell exists. If this argument is omitted, only the row specified by _StartRow_ is included in the range. If this argument is specified but _StartRow_ is omitted, an error occurs.|
|_EndColumn_|Optional| **Long**|The column in which the ending cell exists. If this argument is omitted, only the column specified by _StartColumn_ is included in the range. If this argument is specified but _StartColumn_ is omitted, an error occurs.|

## Remarks

If all arguments are omitted, all the cells in the table are included in the range.


## Example

This example merges the first two cells in the first two rows of the specified table.

```vb
Sub MergeCells() 
 ActiveDocument.Pages(1).Shapes(2).Table _ 
 .Cells(StartRow:=1, StartColumn:=1, _ 
 EndRow:=2, EndColumn:=2).Merge 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]