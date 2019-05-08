---
title: Shape.SetBegin method (Visio)
keywords: vis_sdr.chm11216550
f1_keywords:
- vis_sdr.chm11216550
ms.prod: visio
api_name:
- Visio.Shape.SetBegin
ms.assetid: 257a6ec4-b9c4-4c42-3c57-6e53c1d4d526
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.SetBegin method (Visio)

Moves the begin point of a one-dimensional (1D) shape to the coordinates represented by  _xPos_ and _yPos_.


## Syntax

_expression_. `SetBegin`( `_xPos_` , `_yPos_` )

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The new x-coordinate of the begin point.|
| _yPos_|Required| **Double**|The new y-coordinate of the begin point.|

## Return value

Nothing


## Remarks

The  **SetBegin** method only applies to 1D shapes. If the indicated shape is a 2D shape, an error is generated.

The coordinates represented by the  _xPos_ and _yPos_ arguments are parent coordinates, measured from the origin of the shape's parent (the page or group that contains the shape).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]