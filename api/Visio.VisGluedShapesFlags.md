---
title: VisGluedShapesFlags enumeration (Visio)
keywords: vis_sdr.chm70580
f1_keywords:
- vis_sdr.chm70580
ms.prod: visio
api_name:
- Visio.VisGluedShapesFlags
ms.assetid: c89e043e-3b86-f885-584d-54d20dc5f337
ms.date: 06/08/2017
localization_priority: Normal
---


# VisGluedShapesFlags enumeration (Visio)

Specifies constants that identify which shapes to return, based on the dimensionality and directionality of the connection points that are glued to a particular shape; passed to the  **[Shapes.GluedShapes](Visio.Shape.GluedShapes.md)** method.



|Name|Value|Description|
|:-----|:-----|:-----|
| **visGluedShapesAll1D**|0|Return all 1D shapes that are glued to this shape.|
| **visGluedShapesIncoming1D**|1|Return the 1D shapes whose end points are glued to this shape.|
| **visGluedShapesOutgoing1D**|2|Return the 1D shapes whose begin points are glued to this shape.|
| **visGluedShapesAll2D**|3|Return the 2D shapes that are glued to this shape and the 2D shapes to which this shape is glued. |
| **visGluedShapesIncoming2D**|4|If the source object is a 1D shape, return the 2D shape to which the begin point is glued. If the source object is a 2D shape, return the 2D shapes that are glued to this shape.|
| **visGluedShapesOutgoing2D**|5|If the source object is a 1D shape, return the 2D shape to which the end point is glued. If the source object is a 2D shape, return the 2D shapes to which this shape is glued.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]