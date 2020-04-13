---
title: Font.ColorIndexBi property (Word)
keywords: vbawd10.chm156369060
f1_keywords:
- vbawd10.chm156369060
ms.prod: word
api_name:
- Word.Font.ColorIndexBi
ms.assetid: cadba8bf-8f2d-e9c3-e6f3-af34282bd75c
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.ColorIndexBi property (Word)

Returns or sets the color for the specified  **Font** object in a right-to-left language document. Read/write **WdColorIndex**.


## Syntax

_expression_. `ColorIndexBi`

_expression_ Required. A variable that represents a **[Font](Word.Font.md)** object.


## Remarks

The **wdByAuthor** constant is not valid for **Font** objects.


## Example

This example sets the color of the **Font** object to teal.


```vb
Selection.Font.ColorIndexBi = wdTeal
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]