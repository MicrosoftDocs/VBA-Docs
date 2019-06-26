---
title: Series.ErrorBars property (PowerPoint)
keywords: vbapp10.chm65695
f1_keywords:
- vbapp10.chm65695
ms.prod: powerpoint
api_name:
- PowerPoint.Series.ErrorBars
ms.assetid: 6d3a4bd3-93f1-95d6-6d8e-4f296c1b5f95
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.ErrorBars property (PowerPoint)

Returns the error bars for the series. Read-only  **[ErrorBars](PowerPoint.ErrorBars.md)**.


## Syntax

_expression_. `ErrorBars`

_expression_ A variable that represents a '[Series](PowerPoint.Series.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the error bar color for series one of the first chart in the active document. You should run the example on a 2D line chart that has error bars for series one.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1)

            .ErrorBars.Border.ColorIndex = 8

        End With

    End If

End With


```


## See also


[Series Object](PowerPoint.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]