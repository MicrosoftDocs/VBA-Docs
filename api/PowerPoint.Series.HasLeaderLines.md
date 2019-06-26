---
title: Series.HasLeaderLines property (PowerPoint)
keywords: vbapp10.chm66930
f1_keywords:
- vbapp10.chm66930
ms.prod: powerpoint
api_name:
- PowerPoint.Series.HasLeaderLines
ms.assetid: 4aaab32e-56e7-cd47-c3a2-ff92df218373
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.HasLeaderLines property (PowerPoint)

 **True** if the series has leader lines. Read/write **Boolean**.


## Syntax

_expression_.**HasLeaderLines**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

This property applies only to pie charts.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds data labels and blue leader lines to series one on the pie chart. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .HasDataLabels = True

            .DataLabels.Position = xlLabelPositionBestFit

            .HasLeaderLines = True

            .LeaderLines.Border.ColorIndex = 5

        End With

    End If

End With


```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]