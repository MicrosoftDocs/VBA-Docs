---
title: Document.PresentIt method (Word)
keywords: vbawd10.chm158007551
f1_keywords:
- vbawd10.chm158007551
ms.prod: word
api_name:
- Word.Document.PresentIt
ms.assetid: 2565f8a5-d99d-0b40-aea6-2ad20f9ed07f
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PresentIt method (Word)

Opens PowerPoint with the specified Word document loaded.


## Syntax

_expression_. `PresentIt`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sends a copy of the document named "MyPresentation.doc" to PowerPoint.


```vb
Documents("MyPresentation.doc").PresentIt
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]