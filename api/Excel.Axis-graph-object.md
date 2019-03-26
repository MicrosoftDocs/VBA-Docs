---
title: Axis object (Excel Graph)
keywords: vbagr10.chm5207088
f1_keywords:
- vbagr10.chm5207088
ms.prod: excel
api_name:
- Excel.Axis
ms.assetid: 708d79de-edcc-ac18-58ec-b9921be9b37e
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis object (Excel Graph)

Represents a single axis in a chart. The  **Axis** object is a member of the **[Axes](Excel.Axes-graph-collection.md)** collection.


## Using the Axis Object

Use  **Axes**( _type_,  _group_), where  _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. _Type_ can be one of the following **xlAxisType** constants: **xlCategory**,  **xlSeries**, or  **xlValue**.  _Group_ can be either of the following **xlAxisGroup** constants: **xlPrimary** or **xlSecondary**. For more information, see the  **[Axes](Excel.Axes-graph-method.md)** method.

The following example sets the text of the category axis title in the chart.




```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]