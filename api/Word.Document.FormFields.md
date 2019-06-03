---
title: Document.FormFields property (Word)
keywords: vbawd10.chm158007317
f1_keywords:
- vbawd10.chm158007317
ms.prod: word
api_name:
- Word.Document.FormFields
ms.assetid: ed97fd75-0da5-b008-26c6-ea16465fddc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FormFields property (Word)

Returns a  **[FormFields](Word.formfields.md)** collection that represents all the form fields in the document. Read-only.


## Syntax

_expression_. `FormFields`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the content of the form field named "Text1" to "Name."


```vb
ActiveDocument.FormFields("Text1").Result = "Name"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]