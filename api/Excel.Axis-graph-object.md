---
title: Axis object (Excel Graph)
keywords: vbagr10.chm5207088
f1_keywords:
- vbagr10.chm5207088
ms.prod: excel
api_name:
- Excel.Axis
ms.assetid: 708d79de-edcc-ac18-58ec-b9921be9b37e
ms.date: 04/06/2019
localization_priority: Normal
---


# Axis object (Excel Graph)

Represents a single axis in a chart. The **Axis** object is a member of the **[Axes](Excel.Axes-graph-collection.md)** collection.


## Remarks

Use **[Axes](excel.axes-graph-method.md)** (_type_, _group_), where _type_ is the axis type and _group_ is the axis group, to return a single **Axis** object. 

_Type_ can be one of the following **[XlAxisType](excel.xlaxistype.md)** constants: **xlCategory**, **xlSeriesAxis**, or **xlValue**.  

_Group_ can be either of the following **[XlAxisGroup](excel.xlaxisgroup.md)** constants: **xlPrimary** or **xlSecondary**. 


## Example

The following example sets the text of the category axis title in the chart.

```vb
With myChart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]