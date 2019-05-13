---
title: Shape.AlternativeText property (Excel)
keywords: vbaxl10.chm636132
f1_keywords:
- vbaxl10.chm636132
ms.prod: excel
api_name:
- Excel.Shape.AlternativeText
ms.assetid: 40b53b31-c4e2-0fd8-1a37-fa1e88ccd2be
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.AlternativeText property (Excel)

Returns or sets the descriptive (alternative) text string for a **Shape** object when the object is saved to a webpage. Read/write **String**.


## Syntax

_expression_.**AlternativeText**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Remarks

The alternative text can be displayed either in place of the shape's image in the web browser, or directly over the image when the mouse pointer hovers over the image (in browsers that support these features).


## Example

This example sets the alternative text for the first shape on the first worksheet to a description of the shape.

```vb
Worksheets(1).Shapes(1).AlternativeText = "Concentric circles"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]