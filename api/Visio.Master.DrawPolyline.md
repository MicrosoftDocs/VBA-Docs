---
title: Master.DrawPolyline method (Visio)
keywords: vis_sdr.chm10716215
f1_keywords:
- vis_sdr.chm10716215
ms.prod: visio
api_name:
- Visio.Master.DrawPolyline
ms.assetid: a599e60c-ccd6-ce6b-7e54-f65f8500447d
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.DrawPolyline method (Visio)

Creates a shape whose path is a polyline along a given set of points.


## Syntax

_expression_. `DrawPolyline`( `_xyArray()_` , `_Flags_` )

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _xyArray()_|Required| **Double**| An array of alternating _x_ and _y_ values that defines points in the new shape's path.|
| _Flags_|Required| **Integer**|Flags that influence how the shape is drawn.|

## Return value

Shape


## Remarks

The  **DrawPolyline** method creates a new shape whose path consists of a sequence of line segments and whose endpoints match the points specified in _xyArray_. Calling the **DrawPolyline** method is equivalent to calling the **DrawSpline** method with a tolerance of zero (0) and a flag of **visSplineAbrupt**.

The control points should be in internal drawing units (inches) with respect to the coordinate space of the page, master, or group in which the new shape is being created. The passed array should be a type SAFEARRAY of 8-byte floating point values passed by reference (VT_R8|VT_ARRAY|VT_BYREF). This is how Microsoft Visual Basic passes arrays to Automation objects.

The  _Flags_ argument is a bitmask that specifies options for drawing the new shape. Its value can include **visPolyline1D** (8) or **visPolyarcs** (256). If _Flags_ includes:




-  **visPolyline1D** and if the first and last points in _xyArray_ don't coincide, the **DrawPolyline** method produces a shape with one-dimensional (1D) behavior; otherwise, it produces a shape with two-dimensional (2D) behavior.
    
-  **visPolyarcs**, Microsoft Visio will produce a sequence of arcs rather than a sequence of line segments; _xyArray_ should specify the initial _x,y_ point of the sequence followed by _x,y_ bow triples. Visio will produce a shape with EllipticalArcTo rows where the bow of the arc matches the specified value.
    


If the first and last points in  _xyArray_ coincide, the **DrawPolyline** method produces a filled shape.


## Example

The following example shows how to draw two polyline shapes that have 2D and 1D behavior, respectively, on the active page.


```vb
 
Public Sub DrawPolyline_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim adblXYPoints(1 To 8) As Double 
 Dim intCounter As Integer 
 
 'Initialize array with coordinates. 
 adblXYPoints(1) = 1 
 adblXYPoints(2) = 1 
 adblXYPoints(3) = 3 
 adblXYPoints(4) = 3 
 adblXYPoints(5) = 5 
 adblXYPoints(6) = 1 
 adblXYPoints(7) = 1 
 adblXYPoints(8) = 2 
 
 'Use the DrawPolyline method to draw a shape that has 2D behavior. 
 Set vsoShape = ActivePage.DrawPolyline(adblXYPoints, 0) 
 
 'Increase the Y-coordinate of the array by 4 to separate 
 'the next shape drawn from the first. 
 For intCounter = 2 To UBound(adblXYPoints) Step 2 
 adblXYPoints(intCounter) = adblXYPoints(intCounter) + 4 
 Next intCounter 
 
 'Use the DrawPolyline method to draw a shape that has 1D behavior. 
 Set vsoShape = ActivePage.DrawPolyline(adblXYPoints, visPolyline1D) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]