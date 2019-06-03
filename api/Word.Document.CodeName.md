---
title: Document.CodeName property (Word)
keywords: vbawd10.chm158007558
f1_keywords:
- vbawd10.chm158007558
ms.prod: word
api_name:
- Word.Document.CodeName
ms.assetid: 684f885d-9468-9bc9-d381-ef73286330ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.CodeName property (Word)

Returns the code name for the specified document. Read-only  **String**.


## Syntax

_expression_. `CodeName`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

The code name is the name for the module that houses event macros for a document. The default name for the module is "ThisDocument"; you can view it in the Project window. For information about using events with the Document object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Example

This example returns the name of the code window for the active document.


```vb
Msgbox ActiveDocument.CodeName
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]