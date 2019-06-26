---
title: ChartFont.Size property (PowerPoint)
keywords: vbapp10.chm704010
f1_keywords:
- vbapp10.chm704010
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Size
ms.assetid: 752f7440-3540-5720-5597-b5aa36c52447
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Size property (PowerPoint)

Returns or sets the size of the font. Read/write  **Variant**.


## Syntax

_expression_.**Size**

_expression_ A variable that represents a '[ChartFont](PowerPoint.ChartFont.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font size for the title of the first chart in the active document to 12 points.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Title.Characters.Font.Size = 12

    End If

End With


```


## See also


[ChartFont Object](PowerPoint.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]