---
title: XMLNode.RemoveChild method (Word)
keywords: vbawd10.chm37748838
f1_keywords:
- vbawd10.chm37748838
ms.prod: word
api_name:
- Word.XMLNode.RemoveChild
ms.assetid: 9c4d0e0a-ab58-7c9f-9fc2-f07a28281c29
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNode.RemoveChild method (Word)

Removes a child element from the specified element.


## Syntax

_expression_. `RemoveChild`( `_ChildElement_` )

 _expression_ An expression that returns an [XMLNode](./Word.XMLNode.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChildElement_|Required| **XMLNode**|The child element to be removed.|

## Return value

Nothing


## Example

The following example removes the first child from the first element in the active document.


```vb
ActiveDocument.XMLNodes(1).RemoveChild _ 
 ActiveDocument.XMLNodes(1).ChildNodes(1)
```


## See also


[XMLNode Object](Word.XMLNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]