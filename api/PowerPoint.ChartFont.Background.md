---
title: ChartFont.Background property (PowerPoint)
keywords: vbapp10.chm704001
f1_keywords:
- vbapp10.chm704001
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Background
ms.assetid: 27462713-e2ee-3b2f-ba78-0f29488351b5
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Background property (PowerPoint)

Returns or sets the type of background for text used in charts. Read/write  **Variant** that is set to one of the constants of **[XlBackground](PowerPoint.XlBackground.md)**.


## Syntax

_expression_.**Background**

_expression_ A variable that represents a '[ChartFont](PowerPoint.ChartFont.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a chart title to the first chart in the active document and then sets the font size and specifies a transparent background for the title.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .HasTitle = True

            .ChartTitle.Text = "Rainfall Totals by Month"

            With .ChartTitle.Font

                .Size = 10

                .Background = xlBackgroundTransparent

            End With

        End With

    End If

End With
```


## See also


[ChartFont Object](PowerPoint.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]