---
title: Axis object (Excel)
keywords: vbaxl10.chm560072
f1_keywords:
- vbaxl10.chm560072
ms.prod: excel
api_name:
- Excel.Axis
ms.assetid: 7e08c61b-90f4-8d91-0ee2-84283d10b324
ms.date: 03/29/2019
localization_priority: Normal
---


# Axis object (Excel)

Represents a single axis in a chart.


## Remarks

The **Axis** object is a member of the **[Axes](Excel.Axes(object).md)** collection.

Use **Axes** (_type_, _group_), where _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. 

- _Type_ can be one of the following **[XlAxisType](Excel.XlAxisType.md)** constants: **xlCategory**, **xlSeriesAxis**, or **xlValue**. 

- _Group_ can be one of the following **[XlAxisGroup](Excel.XlAxisGroup.md)** constants: **xlPrimary** or **xlSecondary**. 

For more information, see the **[Axes](Excel.Chart.Axes.md)** method of the **Chart** object.

## Example

The following example sets the category axis title text on the chart sheet named **Chart1**.

```vb
With Charts("chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## Methods

- [Delete](Excel.Axis.Delete.md)
- [Select](Excel.Axis.Select.md)

## Properties

- [Application](Excel.Axis.Application.md)
- [AxisBetweenCategories](Excel.Axis.AxisBetweenCategories.md)
- [AxisGroup](Excel.Axis.AxisGroup.md)
- [AxisTitle](Excel.Axis.AxisTitle.md)
- [BaseUnit](Excel.Axis.BaseUnit.md)
- [BaseUnitIsAuto](Excel.Axis.BaseUnitIsAuto.md)
- [Border](Excel.Axis.Border.md)
- [CategoryNames](Excel.Axis.CategoryNames.md)
- [CategoryType](Excel.Axis.CategoryType.md)
- [Creator](Excel.Axis.Creator.md)
- [Crosses](Excel.Axis.Crosses.md)
- [CrossesAt](Excel.Axis.CrossesAt.md)
- [DisplayUnit](Excel.Axis.DisplayUnit.md)
- [DisplayUnitCustom](Excel.Axis.DisplayUnitCustom.md)
- [DisplayUnitLabel](Excel.Axis.DisplayUnitLabel.md)
- [Format](Excel.Axis.Format.md)
- [HasDisplayUnitLabel](Excel.Axis.HasDisplayUnitLabel.md)
- [HasMajorGridlines](Excel.Axis.HasMajorGridlines.md)
- [HasMinorGridlines](Excel.Axis.HasMinorGridlines.md)
- [HasTitle](Excel.Axis.HasTitle.md)
- [Height](Excel.Axis.Height.md)
- [Left](Excel.Axis.Left.md)
- [LogBase](Excel.Axis.LogBase.md)
- [MajorGridlines](Excel.Axis.MajorGridlines.md)
- [MajorTickMark](Excel.Axis.MajorTickMark.md)
- [MajorUnit](Excel.Axis.MajorUnit.md)
- [MajorUnitIsAuto](Excel.Axis.MajorUnitIsAuto.md)
- [MajorUnitScale](Excel.Axis.MajorUnitScale.md)
- [MaximumScale](Excel.Axis.MaximumScale.md)
- [MaximumScaleIsAuto](Excel.Axis.MaximumScaleIsAuto.md)
- [MinimumScale](Excel.Axis.MinimumScale.md)
- [MinimumScaleIsAuto](Excel.Axis.MinimumScaleIsAuto.md)
- [MinorGridlines](Excel.Axis.MinorGridlines.md)
- [MinorTickMark](Excel.Axis.MinorTickMark.md)
- [MinorUnit](Excel.Axis.MinorUnit.md)
- [MinorUnitIsAuto](Excel.Axis.MinorUnitIsAuto.md)
- [MinorUnitScale](Excel.Axis.MinorUnitScale.md)
- [Parent](Excel.Axis.Parent.md)
- [ReversePlotOrder](Excel.Axis.ReversePlotOrder.md)
- [ScaleType](Excel.Axis.ScaleType.md)
- [TickLabelPosition](Excel.Axis.TickLabelPosition.md)
- [TickLabels](Excel.Axis.TickLabels.md)
- [TickLabelSpacing](Excel.Axis.TickLabelSpacing.md)
- [TickLabelSpacingIsAuto](Excel.Axis.TickLabelSpacingIsAuto.md)
- [TickMarkSpacing](Excel.Axis.TickMarkSpacing.md)
- [Top](Excel.Axis.Top.md)
- [Type](Excel.Axis.Type.md)
- [Width](Excel.Axis.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
