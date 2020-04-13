---
title: Frameset.Width property (Word)
keywords: vbawd10.chm165806083
f1_keywords:
- vbawd10.chm165806083
ms.prod: word
api_name:
- Word.Frameset.Width
ms.assetid: 08c2c81a-119f-18ab-fa6e-5a21ab673cba
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.Width property (Word)

Returns or sets the width (in points) of the specified  **Frameset** object. Read/write **Long**.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

Use the **[WidthType](Word.Frameset.WidthType.md)** property to specify the type of unit in which this value is expressed.


## Example

This example sets the width of the specified  **Frameset** object to 25% of the window width.


```vb
With ActiveWindow.ActivePane.Frameset 
 .WidthType = wdFramesetSizeTypePercent 
 .Width = 25 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]