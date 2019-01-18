---
title: CrossesAt Property
keywords: vbagr10.chm65579
f1_keywords:
- vbagr10.chm65579
ms.prod: excel
api_name:
- Excel.CrossesAt
ms.assetid: aca86ee9-cb90-5982-b1cf-312829d9cc40
ms.date: 06/08/2017
localization_priority: Normal
---


# CrossesAt Property

Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double**.


## Remarks

Setting this property causes the  **[Crosses](Excel.Crosses.md)** property to change to  **xlAxisCrossesCustom**.

This property cannot be used on 3-D charts or radar charts.


## Example

This example sets the category axis in the ActiveChart to cross the value axis at value 3.


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


