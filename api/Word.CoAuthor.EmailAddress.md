---
title: CoAuthor.EmailAddress property (Word)
keywords: vbawd10.chm81068037
f1_keywords:
- vbawd10.chm81068037
ms.prod: word
api_name:
- Word.CoAuthor.EmailAddress
ms.assetid: 48d33e56-78a3-172f-177e-3b250bbec130
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthor.EmailAddress property (Word)

Returns a string that specifies the email address of the specified co author. Read-only.


## Syntax

_expression_. `EmailAddress`

 _expression_ An expression that returns a '[CoAuthor](Word.CoAuthor.md)' object.


## Example

The following code example displays the email address of the first co author in the active document.


```vb
If ActiveDocument.CoAuthoring.Authors.Count <> 0 Then 
 MsgBox ActiveDocument.CoAuthoring.Authors(1).EmailAddress 
Else
 MsgBox "There are no co authors in this document."
End If 
 

```


## See also


[CoAuthor Object](Word.CoAuthor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]