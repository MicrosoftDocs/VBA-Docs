---
title: EndnoteOptions.Location property (Word)
keywords: vbawd10.chm23593060
f1_keywords:
- vbawd10.chm23593060
ms.prod: word
api_name:
- Word.EndnoteOptions.Location
ms.assetid: 3fd348a5-69cd-7319-898e-3f1a102fd644
ms.date: 06/08/2017
localization_priority: Normal
---


# EndnoteOptions.Location property (Word)

Returns or sets the position of all endnotes. Read/write  **WdEndnoteLocation**.


## Syntax

_expression_.**Location** 

_expression_ Required. A variable that represents an '[EndnoteOptions](Word.EndnoteOptions.md)' collection.


## Example

This example positions all endnotes at the end of sections.


```vb
ActiveDocument.Endnotes.Location = wdEndOfSection
```


## See also


[EndnoteOptions Object](Word.EndnoteOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]