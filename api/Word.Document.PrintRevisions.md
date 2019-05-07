---
title: Document.PrintRevisions property (Word)
keywords: vbawd10.chm158007611
f1_keywords:
- vbawd10.chm158007611
ms.prod: word
api_name:
- Word.Document.PrintRevisions
ms.assetid: 2dd7e497-70de-6bd5-7692-5757811fdec7
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PrintRevisions property (Word)

 **True** if revision marks are printed with the document. **False** if revision marks aren't printed (that is, tracked changes are printed as if they'd been accepted). Read/write **Boolean**.


## Syntax

_expression_. `PrintRevisions`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example prints the active document without revision marks.


```vb
With ActiveDocument 
 .PrintRevisions = False 
 .PrintOut 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]