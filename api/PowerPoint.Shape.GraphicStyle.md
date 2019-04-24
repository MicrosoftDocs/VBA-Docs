---
title: Shape.GraphicStyle property (PowerPoint)
keywords: vbapp10.chm547094
f1_keywords:
- vbapp10.chm547094
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.GraphicStyle
ms.date: 04/25/2019
localization_priority: Normal
---


# Shape.GraphicStyle property (PowerPoint)

Returns or sets an **[MsoGraphicStyleIndex](Office.MsoGraphicStyleIndex.md)** constant that represents the style of an SVG graphic. Read/write.

## Syntax

_expression_.**GraphicStyle**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Example

The following code example changes the graphic style for the first shape in the active document.

```vb
Dim myShape As Shape 
 
Set myShape = ActiveDocument.Shapes(1) 
 
myShape.GraphicStyle = msoGraphicStylePreset22
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]