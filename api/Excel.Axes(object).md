---
title: Axes Object (Excel)
keywords: vbaxl10.chm571072
f1_keywords:
- vbaxl10.chm571072
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 581e51e5-3dbb-7f0c-a87d-2d44f67dad0b
ms.date: 06/08/2017
---


# Axes Object (Excel)

A collection of all the  **[Axis](Excel.Axis(object).md)** objects in the specified chart.


## Remarks

Use the  **Axes** method to return the **Axes** collection.

Use  **Axes** ( _type_, _group_ ), where _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](Excel.XlAxisType.md)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _Group_ can be one of the following **[XlAxisGroup](Excel.XlAxisGroup.md)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](Excel.Chart.Axes.md)** method.


## Example

The following example displays the number of axes on embedded chart one on worksheet one.


```vb
With Worksheets(1).ChartObjects(1).Chart 
 MsgBox.Axes.Count 
End With
```

The following example sets the category axis title text on the chart sheet named "Chart1."




```vb
With Charts("chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


## Methods



|**Name**|
|:-----|
|[Item](Excel.Axes.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Axes.Application.md)|
|[Count](Excel.Axes.Count.md)|
|[Creator](Excel.Axes.Creator.md)|
|[Parent](Excel.Axes.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)
