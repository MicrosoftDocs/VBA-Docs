---
title: ChartFont.Underline property (Word)
keywords: vbawd10.chm255918106
f1_keywords:
- vbawd10.chm255918106
ms.prod: word
api_name:
- Word.ChartFont.Underline
ms.assetid: 473bd43d-7f66-b5f1-e9b1-5a37678c332f
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Underline property (Word)

Returns or sets the type of underline applied to the font. Can be one of the  **[XlUnderlineStyle](Word.xlunderlinestyle.md)** constants. Read/write **Variant**.


## Syntax

_expression_.**Underline**

_expression_ A variable that represents a '[ChartFont](Word.ChartFont.md)' object.


## Example

The following example sets the font in the title of the first chart in the active document to single underline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartTitle.Font.Underline = xlUnderlineStyleSingle 
 End If 
End With
```


## See also


[ChartFont Object](Word.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]