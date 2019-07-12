---
title: Axis object (PowerPoint)
keywords: vbapp10.chm682000
f1_keywords:
- vbapp10.chm682000
ms.prod: powerpoint
api_name:
- PowerPoint.Axis
ms.assetid: 38d5e006-ac32-7bdb-f9f0-e8a858dcbf49
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis object (PowerPoint)

Represents a single axis in a chart.


## Remarks

The  **Axis** object is a member of the **[Axes](PowerPoint.Axes.md)** collection.

Use  **Axes** ( _Type_, _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](PowerPoint.XlAxisType.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](PowerPoint.XlAxisGroup.md)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](PowerPoint.Chart.Axes.md)** method.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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



|Name|
|:-----|
|[Delete](PowerPoint.Axis.Delete.md)|
|[Select](PowerPoint.Axis.Select.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Axis.Application.md)|
|[AxisBetweenCategories](PowerPoint.Axis.AxisBetweenCategories.md)|
|[AxisGroup](PowerPoint.Axis.AxisGroup.md)|
|[AxisTitle](PowerPoint.Axis.AxisTitle.md)|
|[BaseUnit](PowerPoint.Axis.BaseUnit.md)|
|[BaseUnitIsAuto](PowerPoint.Axis.BaseUnitIsAuto.md)|
|[Border](PowerPoint.Axis.Border.md)|
|[CategoryNames](PowerPoint.Axis.CategoryNames.md)|
|[CategoryType](PowerPoint.Axis.CategoryType.md)|
|[Creator](PowerPoint.Axis.Creator.md)|
|[Crosses](PowerPoint.Axis.Crosses.md)|
|[CrossesAt](PowerPoint.Axis.CrossesAt.md)|
|[DisplayUnit](PowerPoint.Axis.DisplayUnit.md)|
|[DisplayUnitCustom](PowerPoint.Axis.DisplayUnitCustom.md)|
|[DisplayUnitLabel](PowerPoint.Axis.DisplayUnitLabel.md)|
|[Format](PowerPoint.Axis.Format.md)|
|[HasDisplayUnitLabel](PowerPoint.Axis.HasDisplayUnitLabel.md)|
|[HasMajorGridlines](PowerPoint.Axis.HasMajorGridlines.md)|
|[HasMinorGridlines](PowerPoint.Axis.HasMinorGridlines.md)|
|[HasTitle](PowerPoint.Axis.HasTitle.md)|
|[Height](PowerPoint.Axis.Height.md)|
|[Left](PowerPoint.Axis.Left.md)|
|[LogBase](PowerPoint.Axis.LogBase.md)|
|[MajorGridlines](PowerPoint.Axis.MajorGridlines.md)|
|[MajorTickMark](PowerPoint.Axis.MajorTickMark.md)|
|[MajorUnit](PowerPoint.Axis.MajorUnit.md)|
|[MajorUnitIsAuto](PowerPoint.Axis.MajorUnitIsAuto.md)|
|[MajorUnitScale](PowerPoint.Axis.MajorUnitScale.md)|
|[MaximumScale](PowerPoint.Axis.MaximumScale.md)|
|[MaximumScaleIsAuto](PowerPoint.Axis.MaximumScaleIsAuto.md)|
|[MinimumScale](PowerPoint.Axis.MinimumScale.md)|
|[MinimumScaleIsAuto](PowerPoint.Axis.MinimumScaleIsAuto.md)|
|[MinorGridlines](PowerPoint.Axis.MinorGridlines.md)|
|[MinorTickMark](PowerPoint.Axis.MinorTickMark.md)|
|[MinorUnit](PowerPoint.Axis.MinorUnit.md)|
|[MinorUnitIsAuto](PowerPoint.Axis.MinorUnitIsAuto.md)|
|[MinorUnitScale](PowerPoint.Axis.MinorUnitScale.md)|
|[Parent](PowerPoint.Axis.Parent.md)|
|[ReversePlotOrder](PowerPoint.Axis.ReversePlotOrder.md)|
|[ScaleType](PowerPoint.Axis.ScaleType.md)|
|[TickLabelPosition](PowerPoint.Axis.TickLabelPosition.md)|
|[TickLabels](PowerPoint.Axis.TickLabels.md)|
|[TickLabelSpacing](PowerPoint.Axis.TickLabelSpacing.md)|
|[TickLabelSpacingIsAuto](PowerPoint.Axis.TickLabelSpacingIsAuto.md)|
|[TickMarkSpacing](PowerPoint.Axis.TickMarkSpacing.md)|
|[Top](PowerPoint.Axis.Top.md)|
|[Type](PowerPoint.Axis.Type.md)|
|[Width](PowerPoint.Axis.Width.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]