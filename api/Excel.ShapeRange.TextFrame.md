---
title: ShapeRange.TextFrame property (Excel)
keywords: vbaxl10.chm640097
f1_keywords:
- vbaxl10.chm640097
ms.prod: excel
api_name:
- Excel.ShapeRange.TextFrame
ms.assetid: b72b9c3e-c41c-dce9-46ba-ee156ba52676
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.TextFrame property (Excel)

Returns a **[TextFrame](Excel.TextFrame.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.


## Syntax

_expression_.**TextFrame**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example causes text in the text frame in shape one to be justified. If shape one doesn't have a text frame, this example fails.

```vb
Worksheets(1).Shapes(1).TextFrame _ 
 .HorizontalAlignment = xlHAlignJustify
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]