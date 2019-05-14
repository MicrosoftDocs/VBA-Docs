---
title: Shape.TextFrame2 property (PowerPoint)
keywords: vbapp10.chm547077
f1_keywords:
- vbapp10.chm547077
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.TextFrame2
ms.assetid: bc76d1e5-3feb-51c9-a4d4-61f0bf36183f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.TextFrame2 property (PowerPoint)

Returns the  **[TextFrame2](PowerPoint.TextFrame2.md)** object associated with the specified **[Shape](PowerPoint.Shape.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.


## Syntax

_expression_.**TextFrame2**

 _expression_ An expression that returns a **[Shape](PowerPoint.Shape.md)** object.


## Return value

TextFrame2


## Remarks

Use the  **[TextRange](PowerPoint.TextFrame2.TextRange.md)** property of the **TextFrame2** object to return the text in the text frame.

Use the  **[HasTextFrame](PowerPoint.Shape.HasTextFrame.md)** property to determine whether a shape contains a text frame before you attempt to get the **TextFrame2** property value.


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]