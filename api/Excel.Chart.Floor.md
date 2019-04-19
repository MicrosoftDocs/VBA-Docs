---
title: Chart.Floor property (Excel)
keywords: vbaxl10.chm149109
f1_keywords:
- vbaxl10.chm149109
ms.prod: excel
api_name:
- Excel.Chart.Floor
ms.assetid: 7771ab49-b254-f0f0-a21b-596f541ab6c1
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Floor property (Excel)

Returns a **[Floor](Excel.Floor(object).md)** object that represents the floor of the 3D chart. Read-only.


## Syntax

_expression_.**Floor**

_expression_ An expression that returns a **[Chart](Excel.Chart(object).md)** object.


## Return value

Floor


## Example

This example sets the floor color of Chart1 to blue. The example should be run on a 3D chart (the **Floor** property fails on 2D charts).

```vb
Charts("Chart1").Floor.Interior.ColorIndex = 5
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]