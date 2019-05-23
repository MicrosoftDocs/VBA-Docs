---
title: Selection.ItalicRun method (Word)
keywords: vbawd10.chm158663259
f1_keywords:
- vbawd10.chm158663259
ms.prod: word
api_name:
- Word.Selection.ItalicRun
ms.assetid: 0d36eff1-7308-7695-7058-be79455836ee
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ItalicRun method (Word)

Adds the italic character format to or removes it from the current run.


## Syntax

_expression_. `ItalicRun`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

If the run contains a mix of italic and non-italic text, this method adds the italic character format to the entire run.


## Example

This example toggles the italic formatting for the current selection.


```vb
Selection.ItalicRun
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]