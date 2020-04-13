---
title: Rows.Last property (Word)
keywords: vbawd10.chm155975691
f1_keywords:
- vbawd10.chm155975691
ms.prod: word
api_name:
- Word.Rows.Last
ms.assetid: ae7432c5-6ea8-23eb-6f24-727c79fdd632
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.Last property (Word)

Returns the last item in the **Rows** collection as a **Row** object.


## Syntax

_expression_. `Last`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Example

This example deletes the last row in the first table in the active document.


```vb
ActiveDocument.Tables(1).Rows.Last.Cells.Delete
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]