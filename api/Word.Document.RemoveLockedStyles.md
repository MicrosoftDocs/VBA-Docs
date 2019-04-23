---
title: Document.RemoveLockedStyles method (Word)
keywords: vbawd10.chm158007783
f1_keywords:
- vbawd10.chm158007783
ms.prod: word
api_name:
- Word.Document.RemoveLockedStyles
ms.assetid: 0c20a3c9-b4b3-e9a6-06d1-a9bf9b16dc07
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.RemoveLockedStyles method (Word)

Purges a document of locked styles when formatting restrictions have been applied in a document.


## Syntax

_expression_. `RemoveLockedStyles`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

The following example purges the locked styles in the active document.


```vb
ActiveDocument.RemoveLockedStyles
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]