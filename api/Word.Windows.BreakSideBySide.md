---
title: Windows.BreakSideBySide method (Word)
keywords: vbawd10.chm157351949
f1_keywords:
- vbawd10.chm157351949
ms.prod: word
api_name:
- Word.Windows.BreakSideBySide
ms.assetid: 86e02a0d-4449-30e9-69a1-984e815711d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.BreakSideBySide method (Word)

Ends side by side mode if two windows are in side by side mode. Returns a  **Boolean** that represents whether the method was successful.


## Syntax

_expression_. `BreakSideBySide`

_expression_ Required. A variable that represents a '[Windows](Word.windows.md)' collection.


## Return value

Boolean


## Example

The following example ends side by side mode.


```vb
ActiveDocument.Windows.BreakSideBySide
```


## See also


[Windows Collection Object](Word.windows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]