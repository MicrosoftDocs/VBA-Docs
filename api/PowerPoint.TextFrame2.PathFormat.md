---
title: TextFrame2.PathFormat Property (PowerPoint)
keywords: vbapp10.chm678009
f1_keywords:
- vbapp10.chm678009
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.PathFormat
ms.assetid: 43c83e42-4439-8806-0fbe-688359521426
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.PathFormat Property (PowerPoint)

 Returns or sets the path type for the specified text frame. Read/write.


## Syntax

 _expression_. `PathFormat`

 _expression_ An expression that returns a [TextFrame2](./PowerPoint.TextFrame2.md) object.


## Return value

MsoPathType


## Remarks

The value of the  **PathFormat** property can be one of these **MsoPathType** constants. The value **msoPathTypeMixed** cannot be set. Setting the value **msoPathTypeNone** removes any existing path.


||
|:-----|
|**msoPathType1**|
|**msoPathType2**|
|**msoPathType3**|
|**msoPathType4**|
|**msoPathTypeMixed**|
|**msoPathTypeNone**|

## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]