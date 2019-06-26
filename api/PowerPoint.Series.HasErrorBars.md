---
title: Series.HasErrorBars property (PowerPoint)
keywords: vbapp10.chm65696
f1_keywords:
- vbapp10.chm65696
ms.prod: powerpoint
api_name:
- PowerPoint.Series.HasErrorBars
ms.assetid: 658e45b6-0c1c-af50-491a-d88468782227
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.HasErrorBars property (PowerPoint)

 **True** if the series has error bars. Read/write **Boolean**.


## Syntax

_expression_.**HasErrorBars**

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Remarks

This property is not available for 3D charts. 


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example removes error bars from series one for the first chart in the active document. You should run the example on a 2D line chart that has error bars for series one.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).HasErrorBars = False

    End If

End With
```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]