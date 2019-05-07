---
title: ChartFont.Bold property (Word)
keywords: vbawd10.chm255918082
f1_keywords:
- vbawd10.chm255918082
ms.prod: word
api_name:
- Word.ChartFont.Bold
ms.assetid: e203868f-5234-354e-6488-159351041083
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Bold property (Word)

 **True** if the font is bold. Read/write **Variant**.


## Syntax

_expression_.**Bold**

_expression_ A variable that represents a '[ChartFont](Word.ChartFont.md)' object.


## Example

The following example sets the font to bold for all characters in the chart title of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartTitle.Characters.Font.Bold = True 
 End If 
End With
```


## See also


[ChartFont Object](Word.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]