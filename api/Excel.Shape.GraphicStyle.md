---
title: Shape.GraphicStyle property (Excel)
keywords: vbaxl10.chm636157
f1_keywords:
- vbaxl10.chm636157
ms.prod: excel
api_name:
- Excel.Shape.GraphicStyle
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.GraphicStyle property (Excel)

Returns or sets an **[MsoGraphicStyleIndex](Office.MsoGraphicStyleIndex.md)** constant that represents the style of an SVG graphic. Read/write.


## Syntax

_expression_.**GraphicStyle**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

The following code example changes the graphic style for the first shape in the active document.

```vb
Dim myShape As Shape 
 
Set myShape = ActiveDocument.Shapes(1) 
 
myShape.GraphicStyle = msoGraphicStylePreset22
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]