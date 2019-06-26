---
title: Chart.DepthPercent property (PowerPoint)
keywords: vbapp10.chm684025
f1_keywords:
- vbapp10.chm684025
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.DepthPercent
ms.assetid: f80bbd4f-3a4f-71c0-1859-c71a57aec22b
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DepthPercent property (PowerPoint)

Returns or sets the depth of a 3D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.


## Syntax

_expression_.**DepthPercent**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

This property applies only to 3D charts.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the depth of the first chart in the active document to be 50 percent of its width. You should run this example on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        Chart.DepthPercent = 50

    End If

End With


```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]