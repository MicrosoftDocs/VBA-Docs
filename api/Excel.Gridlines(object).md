---
title: Gridlines object (Excel)
keywords: vbaxl10.chm601072
f1_keywords:
- vbaxl10.chm601072
ms.prod: excel
api_name:
- Excel.Gridlines
ms.assetid: 8a096f01-808f-5708-8da5-5667a5f4080d
ms.date: 03/30/2019
localization_priority: Normal
---


# Gridlines object (Excel)

Represents major or minor gridlines on a chart axis.


## Remarks

Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. This object isn't a collection. There's no object that represents a single gridline; you either have all gridlines for an axis turned on or all of them turned off.

Use the **[MajorGridlines](Excel.Axis.MajorGridlines.md)** property of the **Axis** object to return the **GridLines** object that represents the major gridlines for the axis. Use the **[MinorGridlines](Excel.Axis.MinorGridlines.md)** property to return the **GridLines** object that represents the minor gridlines. It's possible to return both major and minor gridlines at the same time.


## Example

The following example turns on major gridlines for the category axis on the chart sheet named **Chart1**, and then formats the gridlines to be blue dashed lines.

```vb
With Charts("chart1").Axes(xlCategory) 
 .HasMajorGridlines = True 
 .MajorGridlines.Border.Color = RGB(0, 0, 255) 
 .MajorGridlines.Border.LineStyle = xlDash 
End With
```

## Methods

- [Delete](Excel.Gridlines.Delete.md)
- [Select](Excel.Gridlines.Select.md)

## Properties

- [Application](Excel.Gridlines.Application.md)
- [Border](Excel.Gridlines.Border.md)
- [Creator](Excel.Gridlines.Creator.md)
- [Format](Excel.Gridlines.Format.md)
- [Name](Excel.Gridlines.Name.md)
- [Parent](Excel.Gridlines.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
