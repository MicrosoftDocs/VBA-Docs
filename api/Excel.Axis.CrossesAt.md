---
title: Axis.CrossesAt property (Excel)
keywords: vbaxl10.chm561079
f1_keywords:
- vbaxl10.chm561079
ms.prod: excel
api_name:
- Excel.Axis.CrossesAt
ms.assetid: 1cacde6c-567a-d877-9bf1-cec6292e3544
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.CrossesAt property (Excel)

Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double**.


## Syntax

_expression_. `CrossesAt`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Remarks

Setting this property causes the  **[Crosses](Excel.Axis.Crosses.md)** property to change to **xlAxisCrossesCustom**.

This property cannot be used on radar charts. For 3-D charts, this property indicates where the plane defined by the category axes crosses the value axis.


## Example

This example sets the category axis in the active chart to cross the value axis at value 3.


```vb
Sub Chart() 
 
 ' Create a sample source of data. 
 Range("A1") = "2" 
 Range("A2") = "4" 
 Range("A3") = "6" 
 Range("A4") = "3" 
 
 ' Create a chart based on the sample source of data. 
 Charts.Add 
 
 With ActiveChart 
 .ChartType = xlLineMarkersStacked 
 .SetSourceData Source:=Sheets("Sheet1").Range("A1:A4"), PlotBy:= xlColumns 
 .Location Where:=xlLocationAsObject, Name:="Sheet1" 
 End With 
 
 ' Set the category axis to cross the value axis at value 3. 
 ActiveChart.Axes(xlValue).Select 
 Selection.CrossesAt = 3 
 
End Sub
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]