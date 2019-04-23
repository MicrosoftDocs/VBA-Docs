---
title: FootnoteOptions.Location property (Word)
keywords: vbawd10.chm170131556
f1_keywords:
- vbawd10.chm170131556
ms.prod: word
api_name:
- Word.FootnoteOptions.Location
ms.assetid: 29300e96-150f-ea6c-14ce-602816b6907a
ms.date: 06/08/2017
localization_priority: Normal
---


# FootnoteOptions.Location property (Word)

Returns or sets the position of all footnotes. Read/write  **WdFootnoteLocation**.


## Syntax

_expression_.**Location** 

_expression_ Required. A variable that represents a '[Footnote](Word.Footnote.md)' object.


## Example

This example positions footnotes at the bottom of each page.


```vb
ActiveDocument.Footnotes.Location = wdBottomOfPage
```


## See also


[FootnoteOptions Object](Word.FootnoteOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]