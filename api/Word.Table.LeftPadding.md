---
title: Table.LeftPadding property (Word)
keywords: vbawd10.chm156303475
f1_keywords:
- vbawd10.chm156303475
ms.prod: word
api_name:
- Word.Table.LeftPadding
ms.assetid: ad047ad0-7a50-6905-9e60-3a2275e49a62
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.LeftPadding property (Word)

Returns or sets the amount of space (in points) to add to the left of the contents of all the cells in a table. Read/write  **Single**.


## Syntax

_expression_.**LeftPadding**

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

The setting of the  **LeftPadding** property for a single cell overrides the setting of the **LeftPadding** property for the entire table.


## Example

This example sets the left padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).LeftPadding = _ 
 PixelsToPoints(40, False)
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]