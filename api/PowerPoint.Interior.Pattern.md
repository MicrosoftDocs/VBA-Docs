---
title: Interior.Pattern property (PowerPoint)
keywords: vbapp10.chm707004
f1_keywords:
- vbapp10.chm707004
ms.prod: powerpoint
api_name:
- PowerPoint.Interior.Pattern
ms.assetid: f400b457-61ba-e923-debb-14abead41670
ms.date: 06/08/2017
localization_priority: Normal
---


# Interior.Pattern property (PowerPoint)

Returns or sets a  **Variant** value, containing an **[XlPattern](PowerPoint.XlPattern.md)** constant, that represents the interior pattern.


## Syntax

_expression_.**Pattern**

_expression_ A variable that represents an '[Interior](PowerPoint.Interior.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to red, for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasUpDownBars = True

            .DownBars.Interior.Pattern = xlPatternCrissCross

            .DownBars.Interior.PatternColorIndex = 3

        End With

    End If

End With
```


## See also


[Interior Object](PowerPoint.Interior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]