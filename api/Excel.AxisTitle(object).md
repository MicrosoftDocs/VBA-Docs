---
title: AxisTitle Object (Excel)
keywords: vbaxl10.chm564072
f1_keywords:
- vbaxl10.chm564072
ms.prod: excel
api_name:
- Excel.AxisTitle
ms.assetid: 563d3ba5-aa77-b6fc-236a-7838d75eaa53
ms.date: 06/08/2017
---


# AxisTitle Object (Excel)

Represents a chart axis title.


## Remarks

Use the  **[AxisTitle](Excel.Axis.AxisTitle.md)** property to return an **AxisTitle** object.

The  **AxisTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.Axis.HasTitle.md)** property for the axis is **True** .


## Example

The following example activates embedded chart one, sets the value axis title text, sets the font to Bookman 10 point, and formats the word millions as italic.


```vb
Worksheets("sheet1").ChartObjects(1).Activate 
With ActiveChart.Axes(xlValue) 
 .HasTitle = True 
 With .AxisTitle 
 .Caption = "Revenue (millions)" 
 .Font.Name = "bookman" 
 .Font.Size = 10 
 .Characters(10, 8).Font.Italic = True 
 End With 
End With 

```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


