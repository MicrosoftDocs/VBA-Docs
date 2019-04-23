---
title: Legend object (Excel Graph)
keywords: vbagr10.chm131212
f1_keywords:
- vbagr10.chm131212
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: ed529b98-ad11-94b9-68d9-01e325cca58f
ms.date: 04/06/2019
localization_priority: Normal
---


# Legend object (Excel Graph)

Represents the legend in the specified chart. Each chart can have only one legend. 

The **Legend** object contains one or more **[LegendEntry](Excel.LegendEntry-graph-object.md)** objects; each **LegendEntry** object contains a **[LegendKey](Excel.LegendKey-graph-object.md)** object.


## Remarks

Use the **[Legend](excel.legend-graph-property.md)** property to return the **Legend** object. 

The chart legend isn't visible unless the **[HasLegend](Excel.HasLegend.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object fail.


## Example

The following example sets the font style for the legend to bold.

```vb
myChart.Legend.Font.Bold = True
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]