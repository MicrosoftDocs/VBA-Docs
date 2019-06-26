---
title: Document.Unprotect method (Word)
keywords: vbawd10.chm158007417
f1_keywords:
- vbawd10.chm158007417
ms.prod: word
api_name:
- Word.Document.Unprotect
ms.assetid: 04cc2bd3-2af6-de24-bd82-7f489aefdb48
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Unprotect method (Word)

Removes protection from the specified document. .


## Syntax

_expression_.**Unprotect** (_Password_)

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|The password string used to protect the document. Passwords are case-sensitive. If the document is protected with a password and the correct password isn't supplied, a dialog box prompts the user for the password.|

## Security

> [!IMPORTANT] 
> Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security notes for Office solution developers](../Library-Reference/Concepts/security-notes-for-microsoft-office-solution-developers.md). 


## Remarks

If the document isn't protected, this method generates an error.


## Example

This example removes protection from the active document, using the value of the strPassword variable as the password.


```vb
If ActiveDocument.ProtectionType <> wdNoProtection Then 
 ActiveDocument.Unprotect Password:=strPassword 
End If
```

This example removes protection from the active document. Text is inserted, and the document is protected for revisions.




```vb
Set aDoc = ActiveDocument 
If aDoc.ProtectionType <> wdNoProtection Then 
 aDoc.Unprotect 
 Selection.InsertBefore "department six" 
 aDoc.Protect Type:=wdAllowOnlyRevisions, Password:=strPassword 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
