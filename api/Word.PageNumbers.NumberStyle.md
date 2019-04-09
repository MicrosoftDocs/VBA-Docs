---
title: PageNumbers.NumberStyle property (Word)
keywords: vbawd10.chm159776770
f1_keywords:
- vbawd10.chm159776770
ms.prod: word
api_name:
- Word.PageNumbers.NumberStyle
ms.assetid: 5a7a3101-3b16-a107-8790-3666fa7fba54
ms.date: 06/08/2017
localization_priority: Normal
---


# PageNumbers.NumberStyle property (Word)

Returns or sets a  **[WdPageNumberStyle](Word.WdPageNumberStyle.md)** constant that represents the number style. Read/write.


## Syntax

_expression_. `NumberStyle`

_expression_ Required. An expression that returns a '[PageNumbers](Word.pagenumbers.md)' object.


## Example

This example formats the page numbers in the active document's footer as lowercase roman numerals.


```vb
For Each sec In ActiveDocument.Sections 
 sec.Footers(wdHeaderFooterPrimary).PageNumbers _ 
 .NumberStyle = wdPageNumberStyleLowercaseRoman 
Next sec
```


## See also


[PageNumbers Collection Object](Word.pagenumbers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]