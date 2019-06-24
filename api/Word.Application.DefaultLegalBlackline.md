---
title: Application.DefaultLegalBlackline property (Word)
keywords: vbawd10.chm158335435
f1_keywords:
- vbawd10.chm158335435
ms.prod: word
api_name:
- Word.Application.DefaultLegalBlackline
ms.assetid: a22afc29-1f7d-73af-75c2-7ce2fbe2250f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DefaultLegalBlackline property (Word)

 **True** for Microsoft Word to compare and merge documents using the **Legal blackline** option in the **Compare and Merge Documents** dialog box. Read/write **Boolean**.


## Syntax

_expression_. `DefaultLegalBlackline`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example enables Word's Legal blackline option for comparing and merging legal documents.


```vb
Sub CreateLegalBlackline() 
 Application.DefaultLegalBlackline = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]