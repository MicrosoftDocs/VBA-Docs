---
title: Document.Password property (Word)
keywords: vbawd10.chm158007381
f1_keywords:
- vbawd10.chm158007381
ms.prod: word
api_name:
- Word.Document.Password
ms.assetid: 243f1735-5367-4ac9-5643-624ccf501abe
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Password property (Word)

Sets a password that must be supplied to open the specified document. Write-only  **String**.


## Syntax

_expression_. `Password`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

> [!IMPORTANT] 
> Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security notes for Office solution developers](../Library-Reference/Concepts/security-notes-for-microsoft-office-solution-developers.md). 


## Example

This example opens Earnings.doc, sets a password for it, and then closes the document.


```vb
Set myDoc = Documents _ 
 .Open(FileName:="C:\My Documents\Earnings.doc") 
myDoc.Password = strPassword 
myDoc.Close
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]