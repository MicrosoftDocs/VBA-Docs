---
title: Selection.FootnoteOptions property (Word)
keywords: vbawd10.chm158663680
f1_keywords:
- vbawd10.chm158663680
ms.prod: word
api_name:
- Word.Selection.FootnoteOptions
ms.assetid: 064bb3c1-cbaa-9d8f-5b97-a4337b0cfeae
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.FootnoteOptions property (Word)

Returns  **[FootnoteOptions](Word.FootnoteOptions.md)** object that represents the footnotes in a selection.


## Syntax

_expression_. `FootnoteOptions`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example sets the numbering rule in the selection to restart at the beginning of the new section.


```vb
Sub SetFootnoteOptionsRange() 
 Selection.FootnoteOptions.NumberingRule = wdRestartSection 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]