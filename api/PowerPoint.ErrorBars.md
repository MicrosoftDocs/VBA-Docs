---
title: ErrorBars object (PowerPoint)
keywords: vbapp10.chm702000
f1_keywords:
- vbapp10.chm702000
ms.prod: powerpoint
api_name:
- PowerPoint.ErrorBars
ms.assetid: 2c94c8ca-1e27-0f30-5559-788efa301bc0
ms.date: 06/08/2017
localization_priority: Normal
---


# ErrorBars object (PowerPoint)

Represents the error bars on a chart series.


## Remarks

 Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2D chart can have error bars. Only series in scatter groups can have x and y error bars. This object is not a collection. There is no object that represents a single error bar; you either enable x error bars or y error bars for all points in a series or you disable them.

The **[ErrorBar](PowerPoint.Series.ErrorBar.md)** method changes the error bar format and type.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[ErrorBars](PowerPoint.Series.ErrorBars.md)** property to return the **ErrorBars** object. The following example enables error bars for series one of the first chart in the active document and then sets the end style for the error bars.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).HasErrorBars = True

        .Chart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap

    End If

End With
```


## See also



[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]