---
title: Rows.HeightRule property (Word)
keywords: vbawd10.chm155975688
f1_keywords:
- vbawd10.chm155975688
ms.prod: word
api_name:
- Word.Rows.HeightRule
ms.assetid: 478635fd-fcaa-d679-e0e2-b24258615d04
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.HeightRule property (Word)

Returns or sets the rule for determining the height of the specified cells or rows. Read/write  **WdRowHeightRule**.


## Syntax

_expression_. `HeightRule`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Example

This example sets the height rule for the selected rows to automatically adjust to the tallest cell in the row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.HeightRule = wdRowHeightAuto 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]