---
title: PageSetup.PageHeight property (Word)
keywords: vbawd10.chm158400618
f1_keywords:
- vbawd10.chm158400618
ms.prod: word
api_name:
- Word.PageSetup.PageHeight
ms.assetid: f1c557af-65d2-96e6-c796-a9af33dc1730
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.PageHeight property (Word)

Returns or sets the height of the page in points. Read/write  **Single**.


## Syntax

_expression_. `PageHeight`

 _expression_ An expression that returns a **[PageSetup](Word.PageSetup.md)** object.


## Remarks

Setting the  **PageHeight** property changes the **[PaperSize](Word.PageSetup.PaperSize.md)** property to **wdPaperCustom**. Use the **PaperSize** property to set the page height and width to those of a predefined paper size, such as Letter or A4.


## Example

This example sets the page height for the active document to 9 inches.


```vb
With ActiveDocument.PageSetup 
 .PageHeight = InchesToPoints(9) 
 .PageWidth = InchesToPoints(7) 
End With
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]