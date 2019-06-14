---
title: ThreeDFormat.ExtrusionColorType property (Publisher)
keywords: vbapb10.chm3801346
f1_keywords:
- vbapb10.chm3801346
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.ExtrusionColorType
ms.assetid: 5abd895d-0cf3-985d-537e-e45d02f8a852
ms.date: 06/15/2019
localization_priority: Normal
---


# ThreeDFormat.ExtrusionColorType property (Publisher)

Returns or sets an **[MsoExtrusionColorType](Office.MsoExtrusionColorType.md)** constant indicating whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write.


## Syntax

_expression_.**ExtrusionColorType**

_expression_ A variable that represents a **[ThreeDFormat](Publisher.ThreeDFormat.md)** object.


## Return value

MsoExtrusionColorType


## Remarks

The **ExtrusionColorType** property value can be one of the **MsoExtrusionColorType** constants declared in the Microsoft Office type library.


## Example

If the first shape in the active publication has an automatic extrusion color, this example gives the extrusion a custom yellow color.

```vb
With ActiveDocument.Pages(1).Shapes(1).ThreeD 
    If .ExtrusionColorType = msoExtrusionColorAutomatic Then 
        .ExtrusionColor.RGB = RGB(240, 235, 16) 
    End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]