---
title: Cell.Split method (Publisher)
keywords: vbapb10.chm5111844
f1_keywords:
- vbapb10.chm5111844
ms.prod: publisher
api_name:
- Publisher.Cell.Split
ms.assetid: 99bc34df-c8dc-90e5-4262-dbe0a9c9b61d
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.Split method (Publisher)

Splits a merged table cell back into its constituent cells. Returns a **[CellRange](Publisher.CellRange.md)** object representing the constituent cells.


## Syntax

_expression_.**Split**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


## Return value

CellRange


## Remarks

If the specified cell is not a merged cell resulting from using the **[Merge](Publisher.Cell.Merge.md)** method, an error occurs.


## Example

The following example splits the first cell in the table in shape one on page one of the active publication into its constituent cells. Shape one must contain a table, the first cell of which is a merged cell, for this example to work.

```vb
Dim cllMerged As Cell 
 
Set cllMerged = ActiveDocument.Pages(1).Shapes(1).Table.Cells.Item(1) 
 
cllMerged.Split
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]