---
title: Selection.SelectCell method (Word)
keywords: vbawd10.chm158663192
f1_keywords:
- vbawd10.chm158663192
ms.prod: word
api_name:
- Word.Selection.SelectCell
ms.assetid: 49df8e0c-795d-5d5b-79e4-56e0bd64c222
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SelectCell method (Word)

Selects the entire cell containing the current selection.


## Syntax

 _expression_. `SelectCell`

 _expression_ Required. A variable that represents a '[Selection](Word.Selection.md)' object.


## Remarks

To use this method, the current selection must be contained within a single cell.


## Example

This example selects the entire cell containing the current selection.


```vb
Selection.SelectCell
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]