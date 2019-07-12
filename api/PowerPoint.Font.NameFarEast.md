---
title: Font.NameFarEast property (PowerPoint)
keywords: vbapp10.chm575016
f1_keywords:
- vbapp10.chm575016
ms.prod: powerpoint
api_name:
- PowerPoint.Font.NameFarEast
ms.assetid: 0b3f7d98-bda5-eec3-f570-20d8b575c0a3
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.NameFarEast property (PowerPoint)

Returns or sets the Asian font name. Read/write.


## Syntax

_expression_. `NameFarEast`

_expression_ A variable that represents a [Font](PowerPoint.Font.md) object.


## Return value

String


## Remarks

Use the  **[Replace](PowerPoint.Fonts.Replace.md)** method to change the font that's applied to all text and that appears in the **Font** box on the **Font** tab.


## Example

This example displays the name of the Asian font applied to the selection.


```vb
MsgBox ActiveWindow.Selection.ShapeRange _
    .TextFrame.TextRange.Font.NameFarEast
```


## See also


[Font Object](PowerPoint.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]