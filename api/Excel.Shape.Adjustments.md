---
title: Shape.Adjustments property (Excel)
keywords: vbaxl10.chm636089
f1_keywords:
- vbaxl10.chm636089
ms.prod: excel
api_name:
- Excel.Shape.Adjustments
ms.assetid: 425befaf-e058-dff9-2265-66e4f1cbca39
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Adjustments property (Excel)

Returns an **[Adjustments](Excel.Adjustments.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **Shape** object that represents an AutoShape, WordArt, or Connector.


## Syntax

_expression_.**Adjustments**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example sets to 0.25 the value of adjustment one on shape one on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Adjustments(1) = 0.25
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]