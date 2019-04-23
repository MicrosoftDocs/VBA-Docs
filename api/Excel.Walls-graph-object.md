---
title: Walls object (Excel Graph)
keywords: vbagr10.chm5208137
f1_keywords:
- vbagr10.chm5208137
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 97c3a312-abf1-9da7-cbff-8e48737bf499
ms.date: 04/06/2019
localization_priority: Normal
---


# Walls object (Excel Graph)

Represents the walls of the specified 3D chart. 

This object isn't a collection. There's no object that represents a single wall; you must return all the walls as a unit.


## Remarks

Use the **[Walls](excel.walls-graph-property.md)** property to return the **Walls** object. 


## Example

The following example sets the pattern on the walls for the chart. If the chart isn't a 3D chart, this example fails.

```vb
myChart.Walls.Interior.Pattern = xlGray75
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]