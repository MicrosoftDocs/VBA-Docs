---
title: Frameset.Height property (Word)
keywords: vbawd10.chm165806084
f1_keywords:
- vbawd10.chm165806084
ms.prod: word
api_name:
- Word.Frameset.Height
ms.assetid: 4f577980-30ca-540f-932a-a707ab6d8b5f
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.Height property (Word)

Returns or sets a  **Float** that represents the height (in points) of the specified **Frameset** object. Read/write.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

The  **HeightType** property determines the type of unit in which this value is expressed.


## Example

This example sets the height of the specified  **Frameset** object to 25% of the window height.


```vb
With ActiveWindow.ActivePane.Frameset 
 .HeightType = wdFramesetSizeTypePercent 
 .Height = 25 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]