---
title: ShapeRange.HasInkXML property (PowerPoint)
ms.assetid: 1a7b7b8b-c0e8-8f62-1015-e99cb590fd50
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# ShapeRange.HasInkXML property (PowerPoint)

Returns an [MsoTriState](Office.MsoTriState.md) enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the [ShapeRange.InkXML](PowerPoint.shaperange.inkxml.md) property. Read-only.

An error is returned if the shape range does not contain any ink XML.

## Syntax

_expression_. `HasInkXML`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

MsoTriState


## Remarks

The value of the this property can be one of these  **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape range does not contain ink XML.|
|**msoTrue**| The specified shape range does not contain ink XML.|
|**msoTriStateMixed**| The specified shape range contains a mix of msoTrue and msoFalse return values. One or more shapes in the shape range contains ink XML and another shape in the shape range does not contain any ink XML.|

## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)



[MsoTriState](Office.MsoTriState.md)
[ShapeRange.InkXML](PowerPoint.shaperange.inkxml.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]