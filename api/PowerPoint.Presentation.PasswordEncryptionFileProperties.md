---
title: Presentation.PasswordEncryptionFileProperties property (PowerPoint)
keywords: vbapp10.chm583078
f1_keywords:
- vbapp10.chm583078
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PasswordEncryptionFileProperties
ms.assetid: 086ef0bb-5307-1445-3209-f3f79927965c
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.PasswordEncryptionFileProperties property (PowerPoint)

Returns whether Microsoft PowerPoint encrypts file properties for password-protected documents. Read-only.


## Syntax

_expression_. `PasswordEncryptionFileProperties`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

MsoTriState


## Remarks

Use the  **[SetPasswordEncryptionOptions](PowerPoint.Presentation.SetPasswordEncryptionOptions.md)** method to specify the algorithm PowerPoint uses for encrypting documents with passwords.

The value of the  **PasswordEncryptionFileProperties** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|PowerPoint does not encrypt file properties for password-protected documents.|
|**msoTrue**| PowerPoint encrypts file properties for password-protected documents.|

## Example

This example sets the password encryption options if the file properties are not encrypted for password-protected documents.


```vb
Sub PasswordSettings()

    With ActivePresentation
        If .PasswordEncryptionFileProperties = msoFalse Then
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