---
title: Document.PasswordEncryptionKeyLength property (Word)
keywords: vbawd10.chm158007665
f1_keywords:
- vbawd10.chm158007665
ms.prod: word
api_name:
- Word.Document.PasswordEncryptionKeyLength
ms.assetid: 3144a2e8-f787-e38e-4322-66c6e6ac7523
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PasswordEncryptionKeyLength property (Word)

Returns a  **Long** indicating the key length of the algorithm Microsoft Word uses when encrypting documents with passwords. Read-only.


## Syntax

_expression_. `PasswordEncryptionKeyLength`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

Use the **[SetPasswordEncryptionOptions](Word.Document.SetPasswordEncryptionOptions.md)** method to specify the key length Word uses when encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption key length is less than 40.


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionKeyLength < 40 Then 
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