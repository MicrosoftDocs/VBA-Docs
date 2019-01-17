---
title: Master.DrawQuarterArc Method (Visio)
keywords: vis_sdr.chm10751385
f1_keywords:
- vis_sdr.chm10751385
ms.prod: visio
api_name:
- Visio.Master.DrawQuarterArc
ms.assetid: 6c728c0c-8317-6114-70e5-e5cb68a5729f
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.DrawQuarterArc Method (Visio)

Creates a shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.


## Syntax

 _expression_. `DrawQuarterArc`( `_xBegin_` , `_yBegin_` , `_xEnd_` , `_yEnd_` , `_SweepFlag_` )

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _xBegin_|Required| **Double**|The  _x_-coordinate of the begin point of the arc.|
| _yBegin_|Required| **Double**|The  _y_-coordinate of the begin point of the arc.|
| _xEnd_|Required| **Double**|The  _x_-coordinate of the endpoint of the arc.|
| _yEnd_|Required| **Double**|The  _y_-coordinate of the endpoint of the arc.|
| _SweepFlag_|Required| **VisArcSweepFlags**|The type of arc, concave or convex.|

## Return value

Shape


## Remarks

 The begin and endpoints define the bounding rectangle of the arc, and the _SweepFlag_argument determines which of the two possible arcs within the bounding rectangle is drawn. The bounding rectangle is always aligned to the page coordinate system; that is, the  _x_ -axis of the ellipse is parallel to the _x_ -axis of the page.

The following possible values for the  _SweepFlag_ argument are declared in **VisArcSweepFlags** in the Visio type library.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visArcSweepFlagConcave**|0|Concave arc|
| **visArcSweepFlagConvex**|1|Convex arc|

If  _SweepFlag_ is **visArcSweepFlagConcave** , the line joining the center of the ellipse to the arc sweeps through decreasing angles; if it is **visArcSweepFlagConvex** , it sweeps through increasing angles.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DrawQuarterArc** method to draw a concave arc on the drawing page.


```vb
Public Sub DrawQuarterArc_Example 
 
 Dim vsoShape As Visio.Shape 
 Set vsoShape = ActivePage.DrawQuarterArc(3, 3, 6, 8, visArcSweepFlagConcave) 
 
End Sub
```


