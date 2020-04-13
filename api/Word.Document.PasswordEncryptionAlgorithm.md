---
title: Document.PasswordEncryptionAlgorithm property (Word)
keywords: vbawd10.chm158007664
f1_keywords:
- vbawd10.chm158007664
ms.prod: word
api_name:
- Word.Document.PasswordEncryptionAlgorithm
ms.assetid: 5317832f-936b-5c3b-5acc-6c067563acd6
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PasswordEncryptionAlgorithm property (Word)

Returns a  **String** indicating the algorithm Microsoft Word uses for encrypting documents with passwords. Read-only.


## Syntax

_expression_. `PasswordEncryptionAlgorithm`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

Use the **[SetPasswordEncryptionOptions](Word.Document.SetPasswordEncryptionOptions.md)** method to specify the algorithm Word uses for encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption algorithm in use is "OfficeXor," which is the password algorithm used in versions of Word prior to Word 97 for Windows.


```vb
Sub PasswordSettings() 
 With ActiveDocument 
 If .PasswordEncryptionAlgorithm = "OfficeXor" Then 
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