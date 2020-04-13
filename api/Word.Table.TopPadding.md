---
title: Table.TopPadding property (Word)
keywords: vbawd10.chm156303473
f1_keywords:
- vbawd10.chm156303473
ms.prod: word
api_name:
- Word.Table.TopPadding
ms.assetid: 005453cf-019e-c404-3114-c555cf5a1310
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.TopPadding property (Word)

Returns or sets the amount of space (in points) to add above the contents of all the cells in a table. Read/write  **Single**.


## Syntax

_expression_.**TopPadding**

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

The setting of the **TopPadding** property for a single cell overrides the setting of the **TopPadding** property for the entire table.


## Example

This example sets the top padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).TopPadding = _ 
 PixelsToPoints(40, True)
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]