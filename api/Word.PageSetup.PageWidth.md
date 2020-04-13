---
title: PageSetup.PageWidth property (Word)
keywords: vbawd10.chm158400617
f1_keywords:
- vbawd10.chm158400617
ms.prod: word
api_name:
- Word.PageSetup.PageWidth
ms.assetid: 623bf072-b34b-8b8c-a24f-fe6a0f4073ce
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.PageWidth property (Word)

Returns or sets the width of the page in points. Read/write  **Single**.


## Syntax

_expression_. `PageWidth`

 _expression_ An expression that returns a **[PageSetup](Word.PageSetup.md)** object.


## Remarks

Setting the **PageWidth** property changes the **[PaperSize](Word.PageSetup.PaperSize.md)** property to **wdPaperCustom**. Use the **PaperSize** property to set the page height and width to those of a predefined paper size, such as Letter or A4.


## Example

This example returns the page width for Document1. The **[PointsToInches](Word.Global.PointsToInches.md)** method is used to convert points to inches.


```vb
Set doc1set = Documents("Document1").PageSetup 
Msgbox "The page width is " _ 
 & PointsToInches(doc1set.PageWidth) & " inches."
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]