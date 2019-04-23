---
title: Document.Post method (Word)
keywords: vbawd10.chm158007439
f1_keywords:
- vbawd10.chm158007439
ms.prod: word
api_name:
- Word.Document.Post
ms.assetid: 1ff97561-ed82-fcf3-6615-ee7ed27814fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Post method (Word)

Posts the specified document to a public folder in Microsoft Exchange. .


## Syntax

_expression_. `Post`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

This method displays the  **Send to Exchange Folder** dialog box so that a folder can be selected.


## Example

This example displays the  **Send to Exchange Folder** dialog box so that the active document can be posted to a public folder.


```vb
ActiveDocument.Post
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]