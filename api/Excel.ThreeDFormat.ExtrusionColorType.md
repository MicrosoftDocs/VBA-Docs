---
title: ThreeDFormat.ExtrusionColorType property (Excel)
keywords: vbaxl10.chm119007
f1_keywords:
- vbaxl10.chm119007
ms.prod: excel
api_name:
- Excel.ThreeDFormat.ExtrusionColorType
ms.assetid: cb463711-c8a3-5ac4-c81c-83d7b2d6b824
ms.date: 05/17/2019
localization_priority: Normal
---


# ThreeDFormat.ExtrusionColorType property (Excel)

Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write **[MsoExtrusionColorType](Office.MsoExtrusionColorType.md)**.


## Syntax

_expression_.**ExtrusionColorType**

_expression_ A variable that represents a **[ThreeDFormat](Excel.ThreeDFormat.md)** object.


## Example

If shape one on _myDocument_ has an automatic extrusion color, this example gives the extrusion a custom yellow color.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).ThreeD 
    If .ExtrusionColorType = msoExtrusionColorAutomatic Then 
        .ExtrusionColor.RGB = RGB(240, 235, 16) 
    End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]