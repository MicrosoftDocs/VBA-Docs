---
title: ShapeRange.Width Property (PowerPoint)
keywords: vbapp10.chm548042
f1_keywords:
- vbapp10.chm548042
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Width
ms.assetid: b9b521f8-70e0-90aa-fdbf-675c78cc0d28
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Width Property (PowerPoint)

Returns or sets the width of the specified object, in points. Read/write.


## Syntax

 _expression_. `Width`

 _expression_ A variable that represents a [ShapeRange](./PowerPoint.ShapeRange.md) object.


## Return value

Single


## Example

This example arranges windows one and two horizontally; in other words, each window occupies half the available vertical space and all the available horizontal space in the application window's client area. For this example to work, there must be only two document windows open.


```vb
Windows.Arrange ppArrangeTiled

ah = Windows(1).Height                      ' available height

aw = Windows(1).Width + Windows(2).Width    ' available width

With Windows(1)

    .Width = aw

    .Height = ah / 2

    .Left = 0

End With

With Windows(2)

    .Width = aw

    .Height = ah / 2

    .Top = ah / 2

    .Left = 0

End With
```

This example sets the width for column one in the specified table to 80 points (72 points per inch).




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Columns(1).Width = 80
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]