---
title: Document.IsMasterDocument property (Word)
keywords: vbawd10.chm158007342
f1_keywords:
- vbawd10.chm158007342
ms.prod: word
api_name:
- Word.Document.IsMasterDocument
ms.assetid: fadf30e4-9a35-40ef-0b89-ebd981577624
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.IsMasterDocument property (Word)

 **True** if the specified document is a master document. Read-only **Boolean**.


## Syntax

 _expression_. `IsMasterDocument`

 _expression_ A variable that represents a '[Document](Word.Document.md)' object.


## Remarks

A master document includes one or more subdocuments.


## Example

If the active document is a master document, this example switches to master document view and opens the first subdocument.


```vb
If ActiveDocument.IsMasterDocument = True Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 ActiveDocument.Subdocuments(1).Open 
Else 
 MsgBox "This document is not a master document." 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]