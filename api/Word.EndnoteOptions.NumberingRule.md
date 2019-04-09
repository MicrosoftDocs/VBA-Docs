---
title: EndnoteOptions.NumberingRule property (Word)
keywords: vbawd10.chm23593063
f1_keywords:
- vbawd10.chm23593063
ms.prod: word
api_name:
- Word.EndnoteOptions.NumberingRule
ms.assetid: c2690da3-703b-4f9f-cdfb-7ec4e7559b54
ms.date: 06/08/2017
localization_priority: Normal
---


# EndnoteOptions.NumberingRule property (Word)

Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write  **[WdNumberingRule](Word.WdNumberingRule.md)**.


## Syntax

_expression_. `NumberingRule`

_expression_ Required. A variable that represents an '[EndnoteOptions](Word.EndnoteOptions.md)' collection.


## Example

This example restarts endnote numbering after each section break in the active document.


```vb
ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```


## See also


[EndnoteOptions Object](Word.EndnoteOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]