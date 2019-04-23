---
title: Break.PageIndex property (Word)
keywords: vbawd10.chm200343555
f1_keywords:
- vbawd10.chm200343555
ms.prod: word
api_name:
- Word.Break.PageIndex
ms.assetid: cb58716a-801a-11ba-5208-ef8b4e022c97
ms.date: 06/08/2017
localization_priority: Normal
---


# Break.PageIndex property (Word)

Returns a  **Long** that represents the page number on which the specified break occurs.


## Syntax

_expression_. `PageIndex`

 _expression_ An expression that returns a '[Break](Word.Break.md)' object.


## Example

The following code returns the page number on which the specified break occurs.


```vb
ActiveDocument.ActiveWindow.Panes(1).Pages(1).Breaks(1).PageIndex
```


## See also


[Break Object](Word.Break.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]