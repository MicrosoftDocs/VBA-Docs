---
title: Cell.Next property (Word)
keywords: vbawd10.chm156106855
f1_keywords:
- vbawd10.chm156106855
ms.prod: word
api_name:
- Word.Cell.Next
ms.assetid: b4171c7c-6703-9cdf-a964-09e32874fbb6
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Next property (Word)

Returns a **Cell** object that represents the next table cell in the **Cells** collection. Read-only.


## Syntax

_expression_.**Next**

_expression_ A variable that represents a **[Cell](Word.Cell.md)** object.


## Example

If the selection is in a table, this example selects the contents of the next table cell.

```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Next.Select 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]