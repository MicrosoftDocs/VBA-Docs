---
title: XMLNode.ParentNode property (Word)
keywords: vbawd10.chm37748744
f1_keywords:
- vbawd10.chm37748744
ms.prod: word
api_name:
- Word.XMLNode.ParentNode
ms.assetid: 626913c2-d12a-30e3-d1b1-9dd6fb80a30c
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.ParentNode property (Word)

Returns an  **XMLNode** object that represents the parent element of the specified element.


## Syntax

_expression_. `ParentNode`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Example

The following example accesses the parent XML node of the text selected in the active document.


```vb
Dim objNode As XMLNode 
 
Set objNode = Selection.XMLParentNode.ParentNode
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]