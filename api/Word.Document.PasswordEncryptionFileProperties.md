---
title: Document.PasswordEncryptionFileProperties property (Word)
keywords: vbawd10.chm158007666
f1_keywords:
- vbawd10.chm158007666
ms.prod: word
api_name:
- Word.Document.PasswordEncryptionFileProperties
ms.assetid: 8da8be02-636b-bcfb-e12c-14eadf72b3f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PasswordEncryptionFileProperties property (Word)

 **True** if Microsoft Word encrypts file properties for password-protected documents. Read-only **Boolean**.


## Syntax

_expression_. `PasswordEncryptionFileProperties`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

Use the **[SetPasswordEncryptionOptions](Word.Document.SetPasswordEncryptionOptions.md)** method to specify whether Word encrypts file properties for password-protected documents.


## Example

This example sets the password encryption options if the file properties are not encrypted for password-protected documents.


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionFileProperties = False Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
End Sub
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]