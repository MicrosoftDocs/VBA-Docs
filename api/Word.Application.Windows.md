---
title: Application.Windows property (Word)
keywords: vbawd10.chm158334978
f1_keywords:
- vbawd10.chm158334978
ms.prod: word
api_name:
- Word.Application.Windows
ms.assetid: 860d9e12-4c02-be1f-64a7-ef0305881854
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Windows property (Word)

Returns a  **[Windows](Word.windows.md)** collection that represents all document windows. Read-only.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

The collection corresponds to the window names that appear at the bottom of the Window menu. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example arranges all open windows so that they don't overlap.


```vb
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]