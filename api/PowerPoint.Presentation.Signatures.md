---
title: Presentation.Signatures property (PowerPoint)
keywords: vbapp10.chm583067
f1_keywords:
- vbapp10.chm583067
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Signatures
ms.assetid: 978e39bb-298b-d820-63cb-2924bf0770b1
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Signatures property (PowerPoint)

Returns a **SignatureSet** object that represents a collection of digital signatures. Read-only.


## Syntax

_expression_.**Signatures**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

SignatureSet


## Example

The following line of code displays the number of digital signatures.


```vb
Sub DisplayNumberOfSignatures
    MsgBox "Number of digital signatures: " & _
        ActivePresentation.Signatures.Count
End Sub
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]