---
title: Legend.LegendEntries method (PowerPoint)
keywords: vbapp10.chm65709
f1_keywords:
- vbapp10.chm65709
ms.prod: powerpoint
api_name:
- PowerPoint.Legend.LegendEntries
ms.assetid: a6110ddf-76dd-efc9-c6ce-abb260f9534c
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.LegendEntries method (PowerPoint)

Returns a collection of legend entries for the legend.


## Syntax

_expression_. `LegendEntries`

_expression_ A variable that represents a '[Legend](PowerPoint.Legend.md)' object.


## Return value

A **[LegendEntries](PowerPoint.LegendEntries.md)** object that represents the legend entries for the legend.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font for legend entry one on the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.LegendEntries(1).Font.Name = "Arial"

    End If

End With
```


## See also


[Legend Object](PowerPoint.Legend.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]