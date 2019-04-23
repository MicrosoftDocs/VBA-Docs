---
title: Document.Repaginate method (Word)
keywords: vbawd10.chm158007399
f1_keywords:
- vbawd10.chm158007399
ms.prod: word
api_name:
- Word.Document.Repaginate
ms.assetid: 7a45ffbc-6512-6075-69a0-54a9987c27ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Repaginate method (Word)

Repaginates the entire document.


## Syntax

_expression_. `Repaginate`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example repaginates the active document if it has changed since the last time it was saved.


```vb
If ActiveDocument.Saved = False Then ActiveDocument.Repaginate
```

This example repaginates all open documents.




```vb
For Each doc In Documents 
 doc.Repaginate 
Next doc
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]