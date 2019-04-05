---
title: Axes collection (Excel Graph)
keywords: vbagr10.chm131099
f1_keywords:
- vbagr10.chm131099
ms.prod: excel
api_name:
- Excel.Axes
ms.assetid: 89ebeb9d-3c16-0bb0-35a8-9a07483c4eb6
ms.date: 04/06/2019
localization_priority: Normal
---


# Axes collection (Excel Graph)

A collection of all the **[Axis](Excel.Axis-graph-object.md)** objects in the specified chart.


## Remarks

Use the **[Axes](excel.axes-graph-method.md)** method to return the **Axes** collection. 

Use **Axes** (_type_, _group_), where _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object.

_Type_ can be one of the following **[XlAxisType](excel.xlaxistype.md)** constants: **xlCategory**, **xlSeriesAxis**, or **xlValue**.

_Group_ can be either of the following **[XlAxisGroup](excel.xlaxisgroup.md)** constants: **xlPrimary** or **xlSecondary**. 


## Example

The following example displays the number of axes in the chart.

```vb
With myChart 
 MsgBox .Axes.Count 
End With
```

<br/>

The following example sets the title text for the category axis.

```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]