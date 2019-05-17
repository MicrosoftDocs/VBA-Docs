---
title: TextFrame2.WordWrap property (PowerPoint)
keywords: vbapp10.chm678012
f1_keywords:
- vbapp10.chm678012
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.WordWrap
ms.assetid: c087f375-2536-7edf-566d-5934d69fe434
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame2.WordWrap property (PowerPoint)

Determines whether lines of text break automatically to fit inside the shape. Read/write.


## Syntax

_expression_.**WordWrap**

 _expression_ An expression that returns a **[TextFrame2](PowerPoint.TextFrame2.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **WordWrap** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Lines of text do not break to fit within the shape boundaries.|
|**msoTrue**| Lines of text break to fit within the shape boundaries.|

## See also


[TextFrame2 Object](PowerPoint.TextFrame2.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]