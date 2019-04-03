---
title: XMLNode.LastChild property (Word)
keywords: vbawd10.chm37748746
f1_keywords:
- vbawd10.chm37748746
ms.prod: word
api_name:
- Word.XMLNode.LastChild
ms.assetid: 96031a10-c2e9-2ada-67d0-c3c4cad53446
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.LastChild property (Word)

Returns an  **XMLNode** object that represents the last child node of an XML element.


## Syntax

_expression_. `LastChild`

 _expression_ An expression that returns an '[XMLNode](Word.XMLNode.md)' object.


## Example

The following example accesses the last child of the second element in the active document.


```vb
Dim objNode As XMLNode 
 
Set objNode = ActiveDocument.XMLNodes(2).LastChild
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]