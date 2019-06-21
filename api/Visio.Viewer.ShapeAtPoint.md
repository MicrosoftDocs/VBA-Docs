---
title: Viewer.ShapeAtPoint property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ShapeAtPoint
ms.assetid: 0b9562f2-aa9e-5ca2-b3d3-6d6f0f65f5a3
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ShapeAtPoint property (Visio Viewer)

Gets the ID of the shape in the drawing that is open in Microsoft Visio Viewer, at the specified point in the Visio Viewer window, in the coordinate system of the window, measured in pixels. Read-only.


## Syntax

_expression_.**ShapeAtPoint** (_X_, _Y_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_X_|Required| **Long**|The x-coordinate, in pixels, of the point.|
|_Y_|Required| **Long**|The y-coordinate, in pixels, of the point.|

## Return value

**Long**


## Remarks

The origin of the coordinate system of the Visio Viewer window is the upper-left corner. The positive x-direction is to the right, and the positive y-direction is down.

If there is no shape at the specified point, the **ShapeAtPoint** property returns 0.


## Example

The following code gets the ID of the shape at point (200, 200) in the drawing that is open in Visio Viewer.

```vb
Debug.Print vsoViewer.ShapeAtPoint (200, 200)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]