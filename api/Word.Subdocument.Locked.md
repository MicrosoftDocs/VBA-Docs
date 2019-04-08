---
title: Subdocument.Locked property (Word)
keywords: vbawd10.chm159973377
f1_keywords:
- vbawd10.chm159973377
ms.prod: word
api_name:
- Word.Subdocument.Locked
ms.assetid: 787f1a05-48a5-1a37-2eb3-ff2a725e2edd
ms.date: 06/08/2017
localization_priority: Normal
---


# Subdocument.Locked property (Word)

 **True** if a subdocument in a master document is locked. Read/write **Boolean.**


## Syntax

_expression_.**Locked**

_expression_ Required. A variable that represents a '[Subdocument](Word.Subdocument.md)' object.


## Example

This example checks the first subdocument in the specified master document and sets the master document to allow only comments if the subdocument is locked.


```vb
If ActiveDocument.Subdocuments(1).Locked = True Then 
 ActiveDocument.Protect Type:=wdAllowOnlyComments 
End If
```


## See also


[Subdocument Object](Word.Subdocument.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]