---
title: Global.Windows property (Word)
keywords: vbawd10.chm163119106
f1_keywords:
- vbawd10.chm163119106
ms.prod: word
api_name:
- Word.Global.Windows
ms.assetid: 23ebd91a-8f72-4f63-4ad8-95f98e36309c
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Windows property (Word)

Returns a  **Windows** collection that represents all open document windows. Read-only.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example arranges all open windows so that they don't overlap.


```vb
Windows.Arrange ArrangeStyle:=wdTiled
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]