---
title: Corners object (Excel Graph)
keywords: vbagr10.chm131200
f1_keywords:
- vbagr10.chm131200
ms.prod: excel
api_name:
- Excel.Corners
ms.assetid: 2b85affa-f501-5458-67f1-f167bc422507
ms.date: 04/06/2019
localization_priority: Normal
---


# Corners object (Excel Graph)

Represents the corners of the specified 3D chart. This object isn't a collection.


## Remarks

Use the **[Corners](Excel.Corners-graph-property.md)** property to return the **Corners** object. 

If the chart isn't a 3D chart, the **Corners** property fails.

## Example

The following example selects the corners of the chart.

```vb
myChart.Corners.Select
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]