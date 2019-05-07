---
title: ChartFont.Size property (Word)
keywords: vbawd10.chm255918098
f1_keywords:
- vbawd10.chm255918098
ms.prod: word
api_name:
- Word.ChartFont.Size
ms.assetid: 75062920-f306-1bfc-f1e0-e68a19d055e4
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont.Size property (Word)

Returns or sets the size of the font. Read/write  **Variant**.


## Syntax

_expression_.**Size**

_expression_ A variable that represents a '[ChartFont](Word.ChartFont.md)' object.


## Example

The following example sets the font size for the title of the first chart in the active document to 12 points.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Title.Characters.Font.Size = 12 
 End If 
End With 

```


## See also


[ChartFont Object](Word.ChartFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]