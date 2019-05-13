---
title: ShapeRange.TextFrame2 property (PowerPoint)
keywords: vbapp10.chm548086
f1_keywords:
- vbapp10.chm548086
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.TextFrame2
ms.assetid: 56554e58-c16b-09dd-8acc-4e2da7715bef
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.TextFrame2 property (PowerPoint)

Returns the  **[TextFrame2](PowerPoint.TextFrame2.md)** object associated with the specified **[ShapeRange](PowerPoint.ShapeRange.md)** object that contains the alignment and anchoring properties for the specified shape range. Read-only.


## Syntax

_expression_.**TextFrame2**

 _expression_ An expression that returns a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

TextFrame2


## Remarks

Use the  **[TextRange](PowerPoint.TextFrame2.TextRange.md)** property of the **TextFrame2** object to return the text in the text frame.

Use the  **[HasTextFrame](PowerPoint.ShapeRange.HasTextFrame.md)** property to determine whether a shape range contains a text frame before you attempt to get the **TextFrame2** property value.


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]