---
title: ChartFont.FontStyle property (PowerPoint)
keywords: vbapp10.chm704005
f1_keywords:
- vbapp10.chm704005
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.FontStyle
ms.assetid: b93a278e-cf38-ef2a-acdc-862fc4ca0b1c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.FontStyle property (PowerPoint)

Returns or sets the font style. Read/write  **String**.


## Syntax

_expression_.**FontStyle**

_expression_ A variable that represents a '[ChartFont](PowerPoint.ChartFont.md)' object.


## Remarks

Changing this property may affect other  **ChartFont** properties (such as **[Bold](PowerPoint.ChartFont.Bold.md)** and **[Italic](PowerPoint.ChartFont.Italic.md)**).


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font style for the title of the first chart in the active document to bold and italic.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Title.Font.FontStyle = "Bold Italic"

    End If

End With
```


## See also


[ChartFont Object](PowerPoint.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]