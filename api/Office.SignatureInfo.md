---
title: SignatureInfo object (Office)
keywords: vbaof11.chm286000
f1_keywords:
- vbaof11.chm286000
ms.prod: office
api_name:
- Office.SignatureInfo
ms.assetid: fe0ffe7d-7cc7-0d82-6888-d5eacca0d3ce
ms.date: 06/08/2017
localization_priority: Normal
---


# SignatureInfo object (Office)

Represents the information used to create a digital or in-document signature.


## Example

The following example uses the  **GetCertificationDetails** method of the **SignatureInfo** object to get the expiration date of the digital certificate.


```vb
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub 

```


## Methods



|Name|
|:-----|
|[GetCertificateDetail](Office.SignatureInfo.GetCertificateDetail.md)|
|[GetSignatureDetail](Office.SignatureInfo.GetSignatureDetail.md)|
|[SelectCertificateDetailByThumbprint](Office.SignatureInfo.SelectCertificateDetailByThumbprint.md)|
|[SelectSignatureCertificate](Office.SignatureInfo.SelectSignatureCertificate.md)|
|[ShowSignatureCertificate](Office.SignatureInfo.ShowSignatureCertificate.md)|

## Properties



|Name|
|:-----|
|[Application](Office.SignatureInfo.Application.md)|
|[CertificateVerificationResults](Office.SignatureInfo.CertificateVerificationResults.md)|
|[ContentVerificationResults](Office.SignatureInfo.ContentVerificationResults.md)|
|[Creator](Office.SignatureInfo.Creator.md)|
|[IsCertificateExpired](Office.SignatureInfo.IsCertificateExpired.md)|
|[IsCertificateRevoked](Office.SignatureInfo.IsCertificateRevoked.md)|
|[IsCertificateUntrusted](Office.SignatureInfo.IsCertificateUntrusted.md)|
|[IsValid](Office.SignatureInfo.IsValid.md)|
|[ReadOnly](Office.SignatureInfo.ReadOnly.md)|
|[SignatureComment](Office.SignatureInfo.SignatureComment.md)|
|[SignatureImage](Office.SignatureInfo.SignatureImage.md)|
|[SignatureProvider](Office.SignatureInfo.SignatureProvider.md)|
|[SignatureText](Office.SignatureInfo.SignatureText.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
