---
title: ShapeRange.Rotation property (Word)
keywords: vbawd10.chm162857077
f1_keywords:
- vbawd10.chm162857077
ms.prod: word
api_name:
- Word.ShapeRange.Rotation
ms.assetid: c1f28cd0-265c-7d52-e81d-6f242d29779e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Rotation property (Word)

Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write  **Single**.


## Syntax

_expression_.**Rotation**

_expression_ A variable that represents a **[ShapeRange](Word.shaperange.md)** object.


## Remarks

A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. To set the rotation of a three-dimensional shape around the x-axis or the y-axis, use the  **[RotationX](Word.ThreeDFormat.RotationX.md)** property or the **[RotationY](Word.ThreeDFormat.RotationY.md)** property of the **[ThreeDFormat](Word.ThreeDFormat.md)** object.


## See also


[ShapeRange Collection Object](Word.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]