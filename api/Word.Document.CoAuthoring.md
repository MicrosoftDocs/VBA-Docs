---
title: Document.CoAuthoring property (Word)
keywords: vbawd10.chm158007896
f1_keywords:
- vbawd10.chm158007896
ms.prod: word
api_name:
- Word.Document.CoAuthoring
ms.assetid: b67ac270-c583-f141-bf86-6fc385987636
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.CoAuthoring property (Word)

Returns a [CoAuthoring](Word.CoAuthoring.md) object that provides the entry point into the co authoring object model. Read-only.


## Syntax

_expression_. `CoAuthoring`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Remarks

The [CoAuthoring](Word.CoAuthoring.md) object provides information about co authoring at the document level. For example, the [CoAuthoring](Word.CoAuthoring.md) object can provide information about whether there are any locks in the document, which users have current locks on the document, or whether or not updates to the document content is available from the server. Use the **CoAuthoring** property to return the [CoAuthoring](Word.CoAuthoring.md) object.


## Example

The following code example gets a reference to the [CoAuthoring](Word.CoAuthoring.md) object through the **CoAuthoring** property of the active document.


```vb
Dim coAuth As CoAuthoring 
Set coAuth = ActiveDocument.CoAuthoring
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]