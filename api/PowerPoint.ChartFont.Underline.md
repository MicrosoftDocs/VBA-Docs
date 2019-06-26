---
title: ChartFont.Underline property (PowerPoint)
keywords: vbapp10.chm704014
f1_keywords:
- vbapp10.chm704014
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Underline
ms.assetid: b5a3ccf1-97eb-ad6e-6147-2097fd51bf8e
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Underline property (PowerPoint)

Returns or sets the type of underline applied to the font. Can be one of the  **[XlUnderlineStyle](PowerPoint.XlUnderlineStyle.md)** constants. Read/write **Variant**.


## Syntax

_expression_.**Underline**

_expression_ A variable that represents a '[ChartFont](PowerPoint.ChartFont.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font in the title of the first chart in the active document to single underline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartTitle.Font.Underline = xlUnderlineStyleSingle

    End If

End With
```


## See also


[ChartFont Object](PowerPoint.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]