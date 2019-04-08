---
title: Document.Name property (Word)
keywords: vbawd10.chm158007296
f1_keywords:
- vbawd10.chm158007296
ms.prod: word
api_name:
- Word.Document.Name
ms.assetid: 5f5f8938-4dab-19fa-f339-83099c442ec4
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Name property (Word)

Returns the name of the specified object. Read-only  **String**.


## Syntax

_expression_.**Name**

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example returns the name of the first bookmark in Hello.doc.


```vb
abook = Documents("Hello.doc").Bookmarks(1).Name
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]