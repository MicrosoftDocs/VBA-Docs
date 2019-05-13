---
title: Shape.HasChart property (PowerPoint)
keywords: vbapp10.chm547078
f1_keywords:
- vbapp10.chm547078
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.HasChart
ms.assetid: 5de934a4-03c2-959f-b0b9-562217146640
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.HasChart property (PowerPoint)

Returns whether the shape represented by the specified object contains a chart. Read-only.


## Syntax

_expression_.**HasChart**

 _expression_ An expression that returns a **[Shape](PowerPoint.Shape.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **HasChart** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape does not contain a chart.|
|**msoTrue**| The specified shape contains a chart.|

## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]