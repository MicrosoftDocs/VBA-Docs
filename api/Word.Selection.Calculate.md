---
title: Selection.Calculate method (Word)
keywords: vbawd10.chm158662828
f1_keywords:
- vbawd10.chm158662828
ms.prod: word
api_name:
- Word.Selection.Calculate
ms.assetid: a4e7ef08-8442-0579-e738-e4f53ee62d62
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Calculate method (Word)

Calculates a mathematical expression within a selection. Returns the result as a  **Single**.


## Syntax

_expression_. `Calculate`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example calculates the selected mathematical expression and displays the result.


```vb
MsgBox "And the answer is... " & Selection.Calculate
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]