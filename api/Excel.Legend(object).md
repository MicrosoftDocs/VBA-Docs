---
title: Legend object (Excel)
keywords: vbaxl10.chm621072
f1_keywords:
- vbaxl10.chm621072
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: 9be53984-bc9c-f964-9ab3-be52d3699bd9
ms.date: 03/30/2019
localization_priority: Normal
---


# Legend object (Excel)

Represents the legend in a chart. Each chart can have only one legend.


## Remarks

The **Legend** object contains one or more **[LegendEntry](Excel.LegendEntry(object).md)** objects; each **LegendEntry** object contains a **[LegendKey](Excel.LegendKey(object).md)** object.

The chart legend isn't visible unless the **[HasLegend](Excel.Chart.HasLegend.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object will fail.


## Example

Use the **[Legend](Excel.Chart.Legend.md)** property of the **Chart** object to return the **Legend** object. The following example sets the font style for the legend in embedded chart one on worksheet one to bold.

```vb
Worksheets(1).ChartObjects(1).Chart.Legend.Font.Bold = True
```


## Methods

- [Clear](Excel.Legend.Clear.md)
- [Delete](Excel.Legend.Delete.md)
- [LegendEntries](Excel.Legend.LegendEntries.md)
- [Select](Excel.Legend.Select.md)

## Properties

- [Application](Excel.Legend.Application.md)
- [Creator](Excel.Legend.Creator.md)
- [Format](Excel.Legend.Format.md)
- [Height](Excel.Legend.Height.md)
- [IncludeInLayout](Excel.Legend.IncludeInLayout.md)
- [Left](Excel.Legend.Left.md)
- [Name](Excel.Legend.Name.md)
- [Parent](Excel.Legend.Parent.md)
- [Position](Excel.Legend.Position.md)
- [Shadow](Excel.Legend.Shadow.md)
- [Top](Excel.Legend.Top.md)
- [Width](Excel.Legend.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
