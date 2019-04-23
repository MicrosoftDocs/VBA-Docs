---
title: Footnotes.Separator property (Word)
keywords: vbawd10.chm155320424
f1_keywords:
- vbawd10.chm155320424
ms.prod: word
api_name:
- Word.Footnotes.Separator
ms.assetid: 7905cf40-2a04-447e-9cb1-ffdd5fc43bd8
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.Separator property (Word)

Returns a  **[Range](Word.Range.md)** object that represents the footnote separator.


## Syntax

_expression_.**Separator**

_expression_ Required. A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Example

This example changes the footnote separator to a single border indented 3 inches from the right margin.


```vb
With ActiveDocument.Footnotes.Separator 
 .Delete 
 .Borders(wdBorderTop).LineStyle = wdLineStyleSingle 
 .ParagraphFormat.RightIndent = InchesToPoints(3) 
End With
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]