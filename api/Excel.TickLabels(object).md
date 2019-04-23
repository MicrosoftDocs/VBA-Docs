---
title: TickLabels object (Excel)
keywords: vbaxl10.chm615072
f1_keywords:
- vbaxl10.chm615072
ms.prod: excel
api_name:
- Excel.TickLabels
ms.assetid: fcb02bc5-fcdc-db32-168b-2d40e5552991
ms.date: 04/02/2019
localization_priority: Normal
---


# TickLabels object (Excel)

Represents the tick-mark labels associated with tick marks on a chart axis.


## Remarks

This object isn't a collection. There's no object that represents a single tick-mark label; you must return all the tick-mark labels as a unit.

Tick-mark label text for the category axis comes from the name of the associated category in the chart. The default tick-mark label text for the category axis is the number that indicates the position of the category relative to the left end of this axis. To change the number of unlabeled tick marks between tick-mark labels, you must change the **[TickLabelSpacing](Excel.Axis.TickLabelSpacing.md)** property for the category axis.

Tick-mark label text for the value axis is calculated based on the **[MajorUnit](Excel.Axis.MajorUnit.md)**, **[MinimumScale](Excel.Axis.MinimumScale.md)**, and **[MaximumScale](Excel.Axis.MaximumScale.md)** properties of the value axis. To change the tick-mark label text for the value axis, you must change the values of these properties.


## Example

Use the **[TickLabels](Excel.Axis.TickLabels.md)** property of the **Axis** object to return the **TickLabels** object. The following example sets the number format for the tick-mark labels on the value axis in embedded chart one on Sheet1.

```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Axes(xlValue).TickLabels.NumberFormat = "0.00"
```

## Methods

- [Delete](Excel.TickLabels.Delete.md)
- [Select](Excel.TickLabels.Select.md)

## Properties

- [Alignment](Excel.TickLabels.Alignment.md)
- [Application](Excel.TickLabels.Application.md)
- [Creator](Excel.TickLabels.Creator.md)
- [Depth](Excel.TickLabels.Depth.md)
- [Font](Excel.TickLabels.Font.md)
- [Format](Excel.TickLabels.Format.md)
- [MultiLevel](Excel.TickLabels.MultiLevel.md)
- [Name](Excel.TickLabels.Name.md)
- [NumberFormat](Excel.TickLabels.NumberFormat.md)
- [NumberFormatLinked](Excel.TickLabels.NumberFormatLinked.md)
- [NumberFormatLocal](Excel.TickLabels.NumberFormatLocal.md)
- [Offset](Excel.TickLabels.Offset.md)
- [Orientation](Excel.TickLabels.Orientation.md)
- [Parent](Excel.TickLabels.Parent.md)
- [ReadingOrder](Excel.TickLabels.ReadingOrder.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]