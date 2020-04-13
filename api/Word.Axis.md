---
title: Axis object (Word)
keywords: vbawd10.chm1725
f1_keywords:
- vbawd10.chm1725
ms.prod: word
api_name:
- Word.Axis
ms.assetid: 3a7ad7d8-d397-a79a-eb6a-a5f0822cbe5d
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis object (Word)

Represents a single axis in a chart.


## Remarks

The **Axis** object is a member of the **[Axes](Word.Axes.md)** collection.

Use  **Axes** ( _Type_ , _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](Word.xlaxistype.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](Word.xlaxisgroup.md)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](Word.Chart.Axes.md)** method.


## Example

The following example sets the category axis title text for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
 End With 
 End If 
End With
```

## Methods

- [Delete](Word.Axis.Delete.md)
- [Select](Word.Axis.Select.md)

## Properties

- [Application](Word.Axis.Application.md)
- [AxisBetweenCategories](Word.Axis.AxisBetweenCategories.md)
- [AxisGroup](Word.Axis.AxisGroup.md)
- [AxisTitle](Word.Axis.AxisTitle.md)
- [BaseUnit](Word.Axis.BaseUnit.md)
- [BaseUnitIsAuto](Word.Axis.BaseUnitIsAuto.md)
- [Border](Word.Axis.Border.md)
- [CategoryNames](Word.Axis.CategoryNames.md)
- [CategoryType](Word.Axis.CategoryType.md)
- [Creator](Word.Axis.Creator.md)
- [Crosses](Word.Axis.Crosses.md)
- [CrossesAt](Word.Axis.CrossesAt.md)
- [DisplayUnit](Word.Axis.DisplayUnit.md)
- [DisplayUnitCustom](Word.Axis.DisplayUnitCustom.md)
- [DisplayUnitLabel](Word.Axis.DisplayUnitLabel.md)
- [Format](Word.Axis.Format.md)
- [HasDisplayUnitLabel](Word.Axis.HasDisplayUnitLabel.md)
- [HasMajorGridlines](Word.Axis.HasMajorGridlines.md)
- [HasMinorGridlines](Word.Axis.HasMinorGridlines.md)
- [HasTitle](Word.Axis.HasTitle.md)
- [Height](Word.Axis.Height.md)
- [Left](Word.Axis.Left.md)
- [LogBase](Word.Axis.LogBase.md)
- [MajorGridlines](Word.Axis.MajorGridlines.md)
- [MajorTickMark](Word.Axis.MajorTickMark.md)
- [MajorUnit](Word.Axis.MajorUnit.md)
- [MajorUnitIsAuto](Word.Axis.MajorUnitIsAuto.md)
- [MajorUnitScale](Word.Axis.MajorUnitScale.md)
- [MaximumScale](Word.Axis.MaximumScale.md)
- [MaximumScaleIsAuto](Word.Axis.MaximumScaleIsAuto.md)
- [MinimumScale](Word.Axis.MinimumScale.md)
- [MinimumScaleIsAuto](Word.Axis.MinimumScaleIsAuto.md)
- [MinorGridlines](Word.Axis.MinorGridlines.md)
- [MinorTickMark](Word.Axis.MinorTickMark.md)
- [MinorUnit](Word.Axis.MinorUnit.md)
- [MinorUnitIsAuto](Word.Axis.MinorUnitIsAuto.md)
- [MinorUnitScale](Word.Axis.MinorUnitScale.md)
- [Parent](Word.Axis.Parent.md)
- [ReversePlotOrder](Word.Axis.ReversePlotOrder.md)
- [ScaleType](Word.Axis.ScaleType.md)
- [TickLabelPosition](Word.Axis.TickLabelPosition.md)
- [TickLabels](Word.Axis.TickLabels.md)
- [TickLabelSpacing](Word.Axis.TickLabelSpacing.md)
- [TickLabelSpacingIsAuto](Word.Axis.TickLabelSpacingIsAuto.md)
- [TickMarkSpacing](Word.Axis.TickMarkSpacing.md)
- [Top](Word.Axis.Top.md)
- [Type](Word.Axis.Type.md)
- [Width](Word.Axis.Width.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]