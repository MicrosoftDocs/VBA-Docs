---
title: Document.Open Event (Word)
keywords: vbawd10.chm4001005
f1_keywords:
- vbawd10.chm4001005
ms.prod: word
api_name:
- Word.Document.Open
ms.assetid: 80ad090c-69bf-b50e-3171-eab5414309a2
<<<<<<< HEAD
ms.date: 06/08/2017
=======
ms.date: 08/20/2018
>>>>>>> master
---


# Document.Open Event (Word)

Occurs when a document is opened.

<<<<<<< HEAD

## Syntax

Private Sub  _expression_ _'Private Sub Document_Open'

 _expression_ A variable that represents a '[Document](Word.Document.md)' object.
=======
> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.


## Syntax

Private Sub  _expression_ `Private Sub Document_Open`

_expression_ A variable that represents a [Document](Word.Document.md) object.
>>>>>>> master


## Remarks

If the event procedure is stored in a template, the procedure will run when a new document based on that template is opened and when the template itself is opened as a document.

<<<<<<< HEAD
For information about using events with the  **Document** object, see[Using Events with the Document Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).
=======
For information about using events with the **Document** object, see [Using Events with the Document Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).
>>>>>>> master


## Example

<<<<<<< HEAD
This example displays a message when a document is opened. (The procedure can be stored in the  **ThisDocument** class module of a document or its attached template.)
=======
This example displays a message when a document is opened. (The procedure can be stored in the **ThisDocument** class module of a document or its attached template.)
>>>>>>> master


```vb
Private Sub Document_Open() 
 MsgBox "This document is copyrighted." 
End Sub
```


## See also

<<<<<<< HEAD

[Document Object](Word.Document.md)
=======
- [Document Object](Word.Document.md)
>>>>>>> master

