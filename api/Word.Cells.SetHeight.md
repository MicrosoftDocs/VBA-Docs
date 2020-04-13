---
title: Cells.SetHeight method (Word)
keywords: vbawd10.chm155844811
f1_keywords:
- vbawd10.chm155844811
ms.prod: word
api_name:
- Word.Cells.SetHeight
ms.assetid: 116a309b-5687-5fee-e370-a990b310dfcb
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.SetHeight method (Word)

Sets the height of table cells.


## Syntax

_expression_. `SetHeight`( `_RowHeight_` , `_HeightRule_` )

_expression_ Required. A variable that represents a '[Cells](Word.cells.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowHeight_|Required| **Variant**|The height of the row or rows, in points.|
| _HeightRule_|Required| **WdRowHeightRule**|The rule for determining the height of the specified cells.|

## Remarks

Setting the **SetHeight** property of a **Cells** object automatically sets the property for the entire row.


## Example

This example sets the row height of the selected cells to at least 18 points.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells.SetHeight RowHeight:=18, _ 
 HeightRule:=wdRowHeightAtLeast 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]