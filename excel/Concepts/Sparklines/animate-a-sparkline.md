---
title: Animate a Sparkline
ms.prod: excel
ms.assetid: 9a0062c5-4d7a-4236-82c2-7c51fba6f3c9
ms.date: 06/08/2017
localization_priority: Normal
---


# Animate a Sparkline

You can animate a sparkline by using the  [ModifySourceData](../../../api/Excel.SparklineGroup.ModifySourceData.md) method of the [SparklineGroup](../../../api/Excel.SparklineGroup.md) object to iterate over a range of data. This example takes 36 months of data and animates it by displaying the first year of data, then iterates through each subsequent month until it reaches the last month of data. A counter is used to slow the animation so it can be viewed more easily. This example requires a sparkline group that contains three sparklines in the range A2:A4 that represent data in the range B2:AK4.


```vb
Sub SparkAnimation()

    ' The group of sparklines to animate
    Dim oSparkGroup As SparklineGroup
    'variables for the loop
    Dim i As Integer, j As Integer
    
    ' Set up the sparkline group variable
    Set oSparkGroup = Sheet1.Range("A2").SparklineGroups(1)
    
    ' Set the data source to the first year of data
    oSparkGroup.ModifySourceData "B2:M4"
    
    ' Loop through the data points for the subsequent two years
    For i = 1 To 24
        ' Move the reference for the sparkline group over one cell
        oSparkGroup.ModifySourceData Range(oSparkGroup.SourceData).Offset(, 1).Address
        
        ' Slow the animation
        j = 1
        Do
            j = j + 1: DoEvents
        Loop Until j = 4000
    Next i
    
End Sub
```


## See also


 [SparklineGroup Object](../../../api/Excel.SparklineGroup.md)



 <br>
 [Programming With Sparklines In Excel](../../../api/overview/Excel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]