---
title: Cell.PreferredWidthType property (Word)
keywords: vbawd10.chm156106868
f1_keywords:
- vbawd10.chm156106868
ms.prod: word
api_name:
- Word.Cell.PreferredWidthType
ms.assetid: 5880af18-b1a2-cb53-c224-147453e84f0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.PreferredWidthType property (Word)

Returns or sets the preferred unit of measurement to use for the width of the specified cell. Read-only  **WdPreferredWidthType**.


## Syntax

_expression_. `PreferredWidthType`

_expression_ Required. A variable that represents a '[Cell](Word.Cell.md)' object.


## Example

This example sets Microsoft Word to accept widths as a percentage of window width, and then it sets the width of the first table in the document to 50% of the window width.


```vb
With ActiveDocument.Tables(1) 
 .PreferredWidthType = wdPreferredWidthPercent 
 .PreferredWidth = 50 
End With
```


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]