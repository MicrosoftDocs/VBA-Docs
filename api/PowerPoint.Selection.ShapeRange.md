---
title: Selection.ShapeRange property (PowerPoint)
keywords: vbapp10.chm508009
f1_keywords:
- vbapp10.chm508009
ms.prod: powerpoint
api_name:
- PowerPoint.Selection.ShapeRange
ms.assetid: 3fd7aed0-ab63-adaa-1a46-c745b6c3e245
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ShapeRange property (PowerPoint)

Returns a **[ShapeRange](PowerPoint.ShapeRange.md)** object that represents all the slide objects that have been selected on the specified slide. Read-only.


## Syntax

_expression_.**ShapeRange**

_expression_ A variable that represents a [Selection](PowerPoint.Selection.md) object.


## Return value

ShapeRange


## Remarks

The range returned can contain the drawings, shapes, OLE objects, pictures, text objects, titles, headers, footers, slide number placeholder, and date and time objects on a slide.

You can return a shape range from a selection when the presentation is in normal, slide, or any master view.


## Example

This example sets the fill foreground color for all the selected shapes in window one.


```vb
Windows(1).Selection.ShapeRange.Fill _
    .ForeColor.RGB = RGB(255, 0, 255)
```


## See also


[Selection Object](PowerPoint.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]