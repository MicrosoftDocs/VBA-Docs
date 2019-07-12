---
title: Presentation.PasswordEncryptionAlgorithm property (PowerPoint)
keywords: vbapp10.chm583076
f1_keywords:
- vbapp10.chm583076
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PasswordEncryptionAlgorithm
ms.assetid: 728934cf-b4f3-6acd-0e42-6fc5928af807
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.PasswordEncryptionAlgorithm property (PowerPoint)

Returns the algorithm Microsoft PowerPoint uses for encrypting documents with passwords. Read-only.


## Syntax

_expression_. `PasswordEncryptionAlgorithm`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Remarks

Use the  **[SetPasswordEncryptionOptions](PowerPoint.Presentation.SetPasswordEncryptionOptions.md)** method to specify the algorithm PowerPoint uses for encrypting documents with passwords.


## Example

This example sets the password encryption options if the password encryption algorithm in use is not RC4.


```vb
Sub PasswordSettings()
    With ActivePresentation
        If .PasswordEncryptionAlgorithm <> "RC4" Then
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


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]