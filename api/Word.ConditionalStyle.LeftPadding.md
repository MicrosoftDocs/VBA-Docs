---
title: ConditionalStyle.LeftPadding property (Word)
keywords: vbawd10.chm91029509
f1_keywords:
- vbawd10.chm91029509
ms.prod: word
api_name:
- Word.ConditionalStyle.LeftPadding
ms.assetid: 5bb8fdb1-a971-13bc-4977-b0ffdcb95116
ms.date: 06/08/2017
localization_priority: Normal
---


# ConditionalStyle.LeftPadding property (Word)

Returns or sets the amount of space (in points) to add to the left of the contents of a single cell or all the cells in a table. Read/write  **Single**.


## Syntax

_expression_.**LeftPadding**

_expression_ Required. A variable that represents a '[ConditionalStyle](Word.ConditionalStyle.md)' object.


## Remarks

The setting of the  **LeftPadding** property for a single cell overrides the setting of the **LeftPadding** property for the entire table.


## Example

This example sets the left padding for the first table in the active document to 40 pixels.


```vb
ActiveDocument.Tables(1).LeftPadding = _ 
 PixelsToPoints(40, False)
```


## See also


[ConditionalStyle Object](Word.ConditionalStyle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]