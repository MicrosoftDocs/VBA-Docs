---
title: XMLNode.NextSibling property (Word)
keywords: vbawd10.chm37748742
f1_keywords:
- vbawd10.chm37748742
ms.prod: word
api_name:
- Word.XMLNode.NextSibling
ms.assetid: 431dd44b-10cd-f869-a70a-a371d16fef92
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.NextSibling property (Word)

Returns an  **XMLNode** object that represents the next element in the document that is at the same level as the specified element.


## Syntax

_expression_. `NextSibling`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Remarks

If the specified element is the last element in the **XMLNodes** collection, this property returns **Nothing**.


## Example

The following example returns the next sibling element to the second element in the active document.


```vb
Dim objNode As XMLNode 
 
Set objNode = ActiveDocument.XMLNodes(2).NextSibling
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]