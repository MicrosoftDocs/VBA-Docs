---
title: Shape.OnAction property (Excel)
keywords: vbaxl10.chm636120
f1_keywords:
- vbaxl10.chm636120
ms.prod: excel
api_name:
- Excel.Shape.OnAction
ms.assetid: 7b278ba3-75d3-1f97-dbe2-181485a88365
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.OnAction property (Excel)

Returns or sets the name of a macro that's run when the specified object is chosen. Read/write **String**.


## Syntax

_expression_.**OnAction**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Remarks

Setting this property for a menu item overrides any custom help information set up for the menu item with the information set up for the assigned macro.


## Example

This example causes Microsoft Excel to run the ShapeClick procedure whenever shape one is chosen.

```vb
Worksheets(1).Shapes(1).OnAction = "ShapeClick"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
