---
title: Document.Reload method (Word)
keywords: vbawd10.chm158007433
f1_keywords:
- vbawd10.chm158007433
ms.prod: word
api_name:
- Word.Document.Reload
ms.assetid: 4feda9b6-dd7b-2e3c-b822-04684638e9d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Reload method (Word)

Reloads a cached document by resolving the hyperlink to the document and downloading it.


## Syntax

_expression_. `Reload`

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

This method reloads the document asynchronously; that is, statements following the **Reload** method in your procedure may execute before the document is actually reloaded. Because of this, you may get unexpected results from using this method in your macros.


## Example

This example opens and reloads the hyperlink to the address "main" on a local intranet.


```vb
With ActiveDocument 
 .FollowHyperlink Address:="https://main" 
 .Reload 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]