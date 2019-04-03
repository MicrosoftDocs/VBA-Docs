---
title: SignatureInfo object (Office)
keywords: vbaof11.chm286000
f1_keywords:
- vbaof11.chm286000
ms.prod: office
api_name:
- Office.SignatureInfo
ms.assetid: fe0ffe7d-7cc7-0d82-6888-d5eacca0d3ce
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo object (Office)

Represents the information used to create a digital or in-document signature.


## Example

The following example uses the **[GetCertificateDetail](office.signatureinfo.getcertificatedetail.md)** method of the **SignatureInfo** object to get the expiration date of the digital certificate.


```vb
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificateDetail(certdetExpirationDate) 
 
End Sub 

```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]