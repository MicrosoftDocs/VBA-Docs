---
title: SignatureInfo.SelectCertificateDetailByThumbprint method (Office)
keywords: vbaof11.chm286016
f1_keywords:
- vbaof11.chm286016
ms.prod: office
api_name:
- Office.SignatureInfo.SelectCertificateDetailByThumbprint
ms.assetid: 997010ee-330f-433d-c62c-bf211b8351d6
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo.SelectCertificateDetailByThumbprint method (Office)

Displays a dialog box containing information about a digital certificate following verification of the user from a thumbprint.


## Syntax

_expression_.**SelectCertificateDetailByThumbprint**(_bstrThumbprint_)

_expression_ An expression that returns a **[SignatureInfo](Office.SignatureInfo.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrThumbprint_|Required|**String**|Contains information about the signer identified by the thumbprint.|

## Example

The following example displays a dialog box with details about the digital certificate for the user identified from a thumbprint.


```vb
Sub SelectDigCertificate(ByVal strVerificationDetail As String) 
Dim objSignatureInfo As SignatureInfo 
Dim objDialog As Object 
 
objDialog = objSignatureInfo.SelectCertificateDetailByThumbprint(strVerificationDetail) 
 
End Sub 

```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]