---
title: ShapeRange.HasChart property (PowerPoint)
keywords: vbapp10.chm548087
f1_keywords:
- vbapp10.chm548087
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.HasChart
ms.assetid: b863fc82-6f99-d102-a399-fde44af9e877
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.HasChart property (PowerPoint)

Returns whether the shape range represented by the specified object contains a chart. Read-only.


## Syntax

_expression_. `HasChart`

 _expression_ An expression that returns a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **HasChart** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape range does not contain a chart.|
|**msoTrue**| The specified shape range contains a chart.|

## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]