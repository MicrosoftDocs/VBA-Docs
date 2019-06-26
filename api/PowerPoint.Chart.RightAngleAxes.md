---
title: Chart.RightAngleAxes property (PowerPoint)
keywords: vbapp10.chm684040
f1_keywords:
- vbapp10.chm684040
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.RightAngleAxes
ms.assetid: 4bccf442-1cf6-48b9-d67c-5a72561211e0
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.RightAngleAxes property (PowerPoint)

 **True** if the chart axes are at right angles, independent of chart rotation or elevation. Read/write **Boolean**.


## Syntax

_expression_.**RightAngleAxes**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

This property applies only to 3D line, column, and bar charts. 

If this property is set to  **True**, the **[Perspective](PowerPoint.Chart.Perspective.md)** property is ignored.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the axes for the first chart in the active document to intersect at right angles. You should run the example on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.RightAngleAxes = True

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]