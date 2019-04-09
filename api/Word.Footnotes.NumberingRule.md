---
title: Footnotes.NumberingRule property (Word)
keywords: vbawd10.chm155320423
f1_keywords:
- vbawd10.chm155320423
ms.prod: word
api_name:
- Word.Footnotes.NumberingRule
ms.assetid: cae020d6-2071-df40-3537-844a612eed3d
ms.date: 06/08/2017
localization_priority: Normal
---


# Footnotes.NumberingRule property (Word)

Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. Read/write  **[WdNumberingRule](Word.WdNumberingRule.md)**.


## Syntax

_expression_. `NumberingRule`

_expression_ Required. A variable that represents a '[Footnotes](Word.footnotes.md)' collection.


## Example

If the footnote numbering in section one is set to restart after each section break, this example sets the numbering to restart on each page.


```vb
Set myRange = ActiveDocument.Sections(1).Range 
If myRange.Footnotes.NumberingRule = wdRestartSection Then 
 myRange.Footnotes.NumberingRule = wdRestartPage 
End If
```


## See also


[Footnotes Collection Object](Word.footnotes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]