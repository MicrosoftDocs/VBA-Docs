---
title: Cell.Previous property (Word)
keywords: vbawd10.chm156106856
f1_keywords:
- vbawd10.chm156106856
ms.prod: word
api_name:
- Word.Cell.Previous
ms.assetid: 64bc6592-e7ae-15bc-456e-1ba0cb1b2935
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Previous property (Word)

Returns a **Cell** object that represents the previous table cell in the **[Cells](Word.cells.md)** collection. Read-only.


## Syntax

_expression_.**Previous**

_expression_ A variable that represents a **[Cell](Word.Cell.md)** object.


## Example

If the selection is in a table, this example selects the contents of the previous cell.

```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows(1).Cells(3).Previous.Select 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]