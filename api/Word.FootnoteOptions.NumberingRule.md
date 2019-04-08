---
title: FootnoteOptions.NumberingRule property (Word)
keywords: vbawd10.chm170131559
f1_keywords:
- vbawd10.chm170131559
ms.prod: word
api_name:
- Word.FootnoteOptions.NumberingRule
ms.assetid: 40cee00b-0354-cc4c-57d9-86f7df1765dc
ms.date: 06/08/2017
localization_priority: Normal
---


# FootnoteOptions.NumberingRule property (Word)

Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write  **[WdNumberingRule](Word.WdNumberingRule.md)**.


## Syntax

_expression_. `NumberingRule`

_expression_ Required. A variable that represents a '[FootnoteOptions](Word.FootnoteOptions.md)' collection.


## Example

If the footnote numbering in section one is set to restart after each section break, this example sets the numbering to restart on each page.


```vb
Set myRange = ActiveDocument.Sections(1).Range 
If myRange.Footnotes.NumberingRule = wdRestartSection Then 
 myRange.Footnotes.NumberingRule = wdRestartPage 
End If
```


## See also


[FootnoteOptions Object](Word.FootnoteOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]