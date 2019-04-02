---
title: SparklineGroup object (Excel)
keywords: vbaxl10.chm870072
f1_keywords:
- vbaxl10.chm870072
ms.prod: excel
api_name:
- Excel.SparklineGroup
ms.assetid: cc694d97-a3d3-3473-2e37-0ede67b97680
ms.date: 04/02/2019
localization_priority: Normal
---


# SparklineGroup object (Excel)

Represents a group of sparklines.

## Remarks

The **SparklineGroup** object can contain multiple sparklines, and contains the property settings for the group, such as color and axis settings. Each sparkline is represented by a **[Sparkline](Excel.Sparkline.md)** object.

Use the **Modify** method to add or remove sparklines from the sparkline group. Use the **ModifyLocation** method to change the location of the sparkline, and use the **ModifySourceData** method to change the range of the source data.

> [!NOTE] 
> Application.ReferenceStyle must be set to xlA1 to execute SparklineGroups.Add.

## Example

The following code example creates a group of column sparklines at the location A1:A4 that are bound to the source data in the range Sheet2!B1:E4. The series color is changed to display the columns in red.

```vb
Dim mySG As SparklineGroup 
Set mySG = Range("$A$1:$A$4").SparklineGroups.Add(Type:=xlSparkColumn, SourceData:= _ 
 "Sheet2!B1:E4") 
 
mySG.SeriesColor.Color = RGB(255, 0, 0)
```

## Methods

- [Delete](Excel.SparklineGroup.Delete.md)
- [Modify](Excel.SparklineGroup.Modify.md)
- [ModifyDateRange](Excel.SparklineGroup.ModifyDateRange.md)
- [ModifyLocation](Excel.SparklineGroup.ModifyLocation.md)
- [ModifySourceData](Excel.SparklineGroup.ModifySourceData.md)

## Properties

- [Application](Excel.SparklineGroup.Application.md)
- [Axes](Excel.SparklineGroup.Axes.md)
- [Count](Excel.SparklineGroup.Count.md)
- [Creator](Excel.SparklineGroup.Creator.md)
- [DateRange](Excel.SparklineGroup.DateRange.md)
- [DisplayBlanksAs](Excel.sparklinegroup.displayblanksas.md)
- [DisplayHidden](Excel.SparklineGroup.DisplayHidden.md)
- [Item](Excel.SparklineGroup.Item.md)
- [LineWeight](Excel.SparklineGroup.LineWeight.md)
- [Location](Excel.SparklineGroup.Location.md)
- [Parent](Excel.SparklineGroup.Parent.md)
- [PlotBy](Excel.sparklinegroup.plotby.md)
- [Points](Excel.sparklinegroup.points.md)
- [SeriesColor](Excel.SparklineGroup.SeriesColor.md)
- [SourceData](Excel.SparklineGroup.SourceData.md)
- [Type](Excel.SparklineGroup.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]