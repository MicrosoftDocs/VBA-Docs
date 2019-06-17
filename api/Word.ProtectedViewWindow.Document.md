---
title: ProtectedViewWindow.Document property (Word)
keywords: vbawd10.chm231735297
f1_keywords:
- vbawd10.chm231735297
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Document
ms.assetid: a4a3e32e-a697-9d9a-f4ea-a07daa1ea238
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.Document property (Word)

Returns a [Document](Word.Document.md) object associated with the Protected View window. Read-only.


## Syntax

_expression_.**Document**

_expression_ A variable that represents a '[ProtectedViewWindow](Word.ProtectedViewWindow.md)' object.


## Remarks

A document displayed in a Protected View window is not a member of the  **[Documents](Word.Application.Documents.md)** collection. Instead, use the **Document** property to access a document that is displayed in a Protected View window.


## Example

The following code example displays the name of the document in the active Protected View window.


```vb
Dim myDoc As Document 
 
Set myDoc = ActiveProtectedViewWindow.Document 
MsgBox myDoc.Name
```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]