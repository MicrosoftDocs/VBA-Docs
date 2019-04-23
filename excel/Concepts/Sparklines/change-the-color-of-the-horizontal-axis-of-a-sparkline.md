---
title: Change the Color of the Horizontal Axis of a Sparkline
ms.prod: excel
ms.assetid: 46e1bf49-9971-4597-8c03-63b7a6d7c6a1
ms.date: 06/08/2017
localization_priority: Normal
---


# Change the Color of the Horizontal Axis of a Sparkline

You can change the color of the horizontal axis of a sparkline by using the  [Color](../../../api/Excel.SparkColor.Color.md) property of the [SparkColor](../../../api/Excel.SparkColor.md) object. The following code example iterates through three sparkline groups and sets the color of the horizontal axis equal to the fill color in cell A8. This example requires three sparkline groups starting in cells A2, B2, and C2. Cell A8 must be filled with the color that you want to use for the color of the horizontal axis. This example uses [Color](../../../api/Excel.Interior.Color.md) property of the [Interior](../../../api/Excel.Interior(object).md) object to get the color of cell A8.


```vb
Sub AxisColor()
    'The sparkline group
    Dim oSparkGroup As SparklineGroup
    'Loop through the sparkline groups on the sheet
    For Each oSparkGroup In Range("A2:C2").SparklineGroups
        'Show the axis
        oSparkGroup.Axes.Horizontal.Axis.Visible = True
        'Set the color of the axis to the color of cell A8
        oSparkGroup.Axes.Horizontal.Axis.Color.Color = Range("A8").Interior.Color
    Next oSparkGroup
End Sub
```


## See also

- [SparklineGroup Object](../../../api/Excel.SparklineGroup.md)
- [Programming With Sparklines In Excel](../../../api/overview/Excel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]