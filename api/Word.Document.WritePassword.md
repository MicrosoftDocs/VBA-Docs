---
title: Document.WritePassword property (Word)
keywords: vbawd10.chm158007382
f1_keywords:
- vbawd10.chm158007382
ms.prod: word
api_name:
- Word.Document.WritePassword
ms.assetid: e3353e68-1196-d896-d978-2c49ceca2940
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.WritePassword property (Word)

Sets a password for saving changes to the specified document. Write-only **String**.


## Syntax

_expression_.**WritePassword**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

> [!IMPORTANT] 
> Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security notes for Office solution developers](../Library-Reference/Concepts/security-notes-for-microsoft-office-solution-developers.md). 


## Example

If the active document isn't already protected against saving changes, this example sets "secret" as the write password for the document.

```vb
Set myDoc = ActiveDocument 
If myDoc.WriteReserved = False Then myDoc.WritePassword = "secret"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]