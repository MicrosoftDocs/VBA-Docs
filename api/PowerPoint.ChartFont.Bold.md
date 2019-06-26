---
title: ChartFont.Bold property (PowerPoint)
keywords: vbapp10.chm704002
f1_keywords:
- vbapp10.chm704002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Bold
ms.assetid: 5d5a0b2e-5aab-f197-79da-e9bb8d219af9
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Bold property (PowerPoint)

 **True** if the font is bold. Read/write **Variant**.


## Syntax

_expression_.**Bold**

_expression_ A variable that represents a '[ChartFont](PowerPoint.ChartFont.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font to bold for all characters in the chart title of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartTitle.Characters.Font.Bold = True

    End If

End With
```


## See also


[ChartFont Object](PowerPoint.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]