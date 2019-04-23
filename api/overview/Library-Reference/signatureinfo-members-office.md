---
title: SignatureInfo members (Office)
description: Represents the information used to create a digital or in-document signature.
ms.prod: office
ms.assetid: 52c19097-8afb-d35c-a9f7-eae81e91c05d
ms.date: 01/30/2019
localization_priority: Normal
---


# SignatureInfo members (Office)

Represents the information used to create a digital or in-document signature.


## Methods

|Name|Description|
|:-----|:-----|
|[GetCertificateDetail](../../Office.SignatureInfo.GetCertificateDetail.md)|Displays a specified detail related to a digital certificate.|
|[GetSignatureDetail](../../Office.SignatureInfo.GetSignatureDetail.md)|Displays a specified detail related to a signature.|
|[SelectCertificateDetailByThumbprint](../../Office.SignatureInfo.SelectCertificateDetailByThumbprint.md)|Displays a dialog box containing information about a digital certificate following verification of the user from a thumbprint.|
|[SelectSignatureCertificate](../../Office.SignatureInfo.SelectSignatureCertificate.md)|Displays a dialog box that allows users to select which signature certificate to use for signing a document.|
|[ShowSignatureCertificate](../../Office.SignatureInfo.ShowSignatureCertificate.md)|Displays the selected or default digital certificate. |

## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.SignatureInfo.Application.md)|Gets an **Application** object that represents the container application for the **SignatureInfo** object. Read-only.|
|[CertificateVerificationResults](../../Office.SignatureInfo.CertificateVerificationResults.md)|Gets the results from the verification of a digital certificate. Read-only.|
|[ContentVerificationResults](../../Office.SignatureInfo.ContentVerificationResults.md)|Gets a value representing the results of the verification of the hashed contents of a signed document. Read-only.|
|[Creator](../../Office.SignatureInfo.Creator.md)|Gets a 32-bit integer that indicates the application in which the **SignatureInfo** object was created. Read-only.|
|[IsCertificateExpired](../../Office.SignatureInfo.IsCertificateExpired.md)|Gets a **Boolean** value indicating whether the digital certificate is expired. Read-only.|
|[IsCertificateRevoked](../../Office.SignatureInfo.IsCertificateRevoked.md)|Gets a **Boolean** value indicating whether the digital certificate is revoked. Read-only.|
|[IsCertificateUntrusted](../../Office.SignatureInfo.IsCertificateUntrusted.md)|Gets a **Boolean** value indicating whether the digital certificate used to digitally sign a document comes from a trusted source. Read-only.|
|[IsValid](../../Office.SignatureInfo.IsValid.md)|Gets a **Boolean** value indicating whether the signature was successfully validated following signature verification. Read-only.|
|[ReadOnly](../../Office.SignatureInfo.ReadOnly.md)|Gets a **Boolean** value indicating whether the **SignatureInfo** object is read-only. Read-only.|
|[SignatureComment](../../Office.SignatureInfo.SignatureComment.md)|Gets or sets a value containing comments included in a signature packet. Read/write.|
|[SignatureImage](../../Office.SignatureInfo.SignatureImage.md)|Gets or sets the value of the image used to sign the document. Read/write.|
|[SignatureProvider](../../Office.SignatureInfo.SignatureProvider.md)|Gets a value identifying an installed signature provider add-in. Read-only.|
|[SignatureText](../../Office.SignatureInfo.SignatureText.md)|Gets or sets the value of the signature text used to sign this document. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]