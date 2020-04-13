---
title: XMLNode.PreviousSibling property (Word)
keywords: vbawd10.chm37748743
f1_keywords:
- vbawd10.chm37748743
ms.prod: word
api_name:
- Word.XMLNode.PreviousSibling
ms.assetid: f4935228-6aaf-e763-27eb-71ed25c1eb6a
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.PreviousSibling property (Word)

Returns an  **XMLNode** object that represents the previous element in the document that is at the same level as the specified element.


## Syntax

_expression_. `PreviousSibling`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Remarks

If the specified element is the first element in the **XMLNodes** collection, this property returns **Nothing**.


## Example

The following example returns the previous sibling element to the third element in the active document.


```vb
Dim objNode As XMLNode 
 
Set objNode = ActiveDocument.XMLNodes(3).PreviousSibling
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]