---
title: Document.Open event (Word)
keywords: vbawd10.chm4001005
f1_keywords:
- vbawd10.chm4001005
ms.prod: word
api_name:
- Word.Document.Open
ms.assetid: 80ad090c-69bf-b50e-3171-eab5414309a2
ms.date: 08/20/2018
localization_priority: Normal
---


# Document.Open event (Word)

Occurs when a document is opened.

> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.


## Syntax

Private Sub  _expression_ `Private Sub Document_Open`

_expression_ A variable that represents a [Document](Word.Document.md) object.


## Remarks

If the event procedure is stored in a template, the procedure will run when a new document based on that template is opened and when the template itself is opened as a document.

For information about using events with the **Document** object, see [Using events with the Document object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Example

This example displays a message when a document is opened. (The procedure can be stored in the **ThisDocument** class module of a document or its attached template.)


```vb
Private Sub Document_Open() 
 MsgBox "This document is copyrighted." 
End Sub
```


## See also

- [Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
