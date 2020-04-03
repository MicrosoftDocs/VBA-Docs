---
title: Shape.Top property (PowerPoint)
keywords: vbapp10.chm547037
f1_keywords:
- vbapp10.chm547037
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Top
ms.assetid: cf56f128-43d7-4f6e-f34c-83fbae854c12
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Top property (PowerPoint)

Returns or sets a  **Single** that represents the distance from the top edge of the shape's bounding box to the top edge of the document. Read/write.


## Syntax

_expression_.**Top**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

Single


## Example

This example arranges windows one and two horizontally; in other words, each window occupies half the available vertical space and all the available horizontal space in the application window's client area. For this example to work, there must be only two document windows open.


```vb
Windows.Arrange ppArrangeTiled

sngHeight = Windows(1).Height                     ' available height

sngWidth = Windows(1).Width + Windows(2).Width    ' available width

With Windows(1)

    .Width = sngWidth

    .Height = sngHeight / 2

    .Left = 0

End With

With Windows(2)

    .Width = sngWidth

    .Height = sngHeight / 2

    .Top = sngHeight / 2

    .Left = 0

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]