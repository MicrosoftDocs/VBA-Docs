---
title: DropCap.Enable method (Word)
keywords: vbawd10.chm156631141
f1_keywords:
- vbawd10.chm156631141
ms.prod: word
api_name:
- Word.DropCap.Enable
ms.assetid: 7e4bdd80-696c-c225-8f7e-0debdf071f27
ms.date: 06/08/2017
localization_priority: Normal
---


# DropCap.Enable method (Word)

Formats the first character in the specified paragraph as a dropped capital letter.


## Syntax

_expression_. `Enable`

_expression_ Required. A variable that represents a '[DropCap](Word.DropCap.md)' object.


## Example

This example formats the first paragraph in the selection to begin with a dropped capital letter.


```vb
With Selection.Paragraphs(1).DropCap 
 .Enable 
 .LinesToDrop = 2 
 .FontName = "Arial" 
End With
```


## See also


[DropCap Object](Word.DropCap.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]