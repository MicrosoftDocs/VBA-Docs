---
title: Table.Spacing property (Word)
keywords: vbawd10.chm156303477
f1_keywords:
- vbawd10.chm156303477
ms.prod: word
api_name:
- Word.Table.Spacing
ms.assetid: 56444e6f-70b6-c815-9098-e6e3ac2d6c3b
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Spacing property (Word)

Returns or sets the spacing (in points) between the cells in a table. Read/write  **Single**.


## Syntax

_expression_. `Spacing`

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Example

This example sets the spacing between cells in the first table in the active document to nine points.


```vb
ActiveDocument.Tables(1).Spacing = 9
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]