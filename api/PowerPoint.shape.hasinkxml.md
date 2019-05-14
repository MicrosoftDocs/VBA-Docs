---
title: Shape.HasInkXML property (PowerPoint)
ms.assetid: 3d985f9b-64e3-8712-fd5f-73d38ca56810
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Shape.HasInkXML property (PowerPoint)

Returns an [MsoTriState](Office.MsoTriState.md) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the [Shape.InkXML](PowerPoint.shape.inkxml.md) property. Read-only.

An error is returned if the shape does not contain any ink XML.

## Syntax

_expression_. `HasInkXML`

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

MsoTriState


## Remarks

The value of this property can be one of these  **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape does not contain ink XML.|
|**msoTrue**| The specified shape contains ink XML.|

## See also


[Shape Object](PowerPoint.Shape.md)



[MsoTriState](Office.MsoTriState.md)
[Shape.InkXML](PowerPoint.shape.inkxml.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]