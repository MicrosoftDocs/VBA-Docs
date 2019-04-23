---
title: SignatureInfo.GetCertificateDetail method (Office)
keywords: vbaof11.chm286007
f1_keywords:
- vbaof11.chm286007
ms.prod: office
api_name:
- Office.SignatureInfo.GetCertificateDetail
ms.assetid: f3cab134-5560-be37-25b4-2cbbfcf0693e
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo.GetCertificateDetail method (Office)

Displays a specified detail related to a digital certificate.


## Syntax

_expression_.**GetCertificateDetail**(_certdet_)

_expression_ An expression that returns a **[SignatureInfo](Office.SignatureInfo.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _certdet_|Required|**CertificateDetail**|An enumerated value specifying which certificate detail to display.|

## Return value

Variant


## Example

The following example gets the expiration date of the digital certificate.


```vb
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificateDetail(certdetExpirationDate) 
 
End Sub 

```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]