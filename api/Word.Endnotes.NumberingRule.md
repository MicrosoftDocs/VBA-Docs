---
title: Endnotes.NumberingRule property (Word)
keywords: vbawd10.chm155254887
f1_keywords:
- vbawd10.chm155254887
ms.prod: word
api_name:
- Word.Endnotes.NumberingRule
ms.assetid: 8f21cc55-b065-86fc-0bc5-d54e9f0e58ac
ms.date: 06/08/2017
localization_priority: Normal
---


# Endnotes.NumberingRule property (Word)

Returns or sets the way endnotes are numbered after page breaks or section breaks. Read/write  **[WdNumberingRule](Word.WdNumberingRule.md)**.


## Syntax

_expression_. `NumberingRule`

_expression_ Required. A variable that represents an '[Endnotes](Word.endnotes.md)' collection.


## Example

This example restarts endnote numbering after each section break in the active document.


```vb
ActiveDocument.Endnotes.NumberingRule = wdRestartSection
```


## See also


[Endnotes Collection Object](Word.endnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]