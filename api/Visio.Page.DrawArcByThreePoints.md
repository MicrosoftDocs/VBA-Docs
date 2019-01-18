---
title: Page.DrawArcByThreePoints Method (Visio)
keywords: vis_sdr.chm10951415
f1_keywords:
- vis_sdr.chm10951415
ms.prod: visio
api_name:
- Visio.Page.DrawArcByThreePoints
ms.assetid: dfa20dfd-22f7-6d99-2649-d8401bf93a19
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.DrawArcByThreePoints Method (Visio)

Creates a shape whose path consists of an arc defined by the three points passed as parameters.


## Syntax

 _expression_. `DrawArcByThreePoints`( `_xBegin_` , `_yBegin_` , `_xEnd_` , `_yEnd_` , `_xControl_` , `_yControl_` )

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _xBegin_|Required| **Double**|The x-coordinate of the begin point of the arc.|
| _yBegin_|Required| **Double**|The y-coordinate of the begin point of the arc.|
| _xEnd_|Required| **Double**|The x-coordinate of the endpoint of the arc.|
| _yEnd_|Required| **Double**|The y-coordinate of the endpoint of the arc.|
| _xControl_|Required| **Double**|The x-coordinate of the control point of the arc.|
| _yControl_|Required| **Double**|The y-coordinate of the control point of the arc.|

## Return value

Shape


## Remarks

All points should be in internal drawing units with respect to the coordinate space of the master, page, or group where the shape is being drawn.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DrawArcByThreePoints** method to draw an arc on the drawing page.


```vb
Public Sub DrawArcByThreePoints_Example 
 
 Dim vsoShape As Visio.Shape 
 Set vsoShape = ActivePage.DrawArcByThreePoints(3, 3, 6, 8, 5, 5) 
 
End Sub
```


