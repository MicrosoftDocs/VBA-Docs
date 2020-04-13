---
title: Table.BottomPadding property (Word)
keywords: vbawd10.chm156303474
f1_keywords:
- vbawd10.chm156303474
ms.prod: word
api_name:
- Word.Table.BottomPadding
ms.assetid: d4e37a85-d194-8d19-c43f-09d30187e007
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.BottomPadding property (Word)

Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write  **Single**.


## Syntax

_expression_.**BottomPadding**

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Remarks

The setting of the **BottomPadding** property for a single cell overrides the setting of the **BottomPadding** property for the entire table.


## Example

This example sets the bottom padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).BottomPadding = _ 
 PixelsToPoints(40, True)
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]