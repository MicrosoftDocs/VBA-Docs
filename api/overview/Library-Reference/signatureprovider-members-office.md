---
title: SignatureProvider members (Office)
description: Represents a signature provider add-in.
ms.prod: office
ms.assetid: 8f99b46b-ee6c-54eb-570a-d2b34c0a8b3d
ms.date: 01/30/2019
localization_priority: Normal
---


# SignatureProvider members (Office)

Represents a signature provider add-in.


## Methods

|Name|Description|
|:-----|:-----|
|[GenerateSignatureLineImage](../../Office.SignatureProvider.GenerateSignatureLineImage.md)|Gets a signature line image.|
|[GetProviderDetail](../../Office.SignatureProvider.GetProviderDetail.md)|Queries the signature provider add-in for various details. |
|[HashStream](../../Office.SignatureProvider.HashStream.md)|Allows a signature provider add-in to create a hash value for the document that you can use to determine if the document contents were tampered with after digital signing.|
|[NotifySignatureAdded](../../Office.SignatureProvider.NotifySignatureAdded.md)|Used to display a dialog box informing the user that the signing process has completed and providing additional functionality for the add-in.|
|[ShowSignatureDetails](../../Office.SignatureProvider.ShowSignatureDetails.md)|Provides a signature provider add-in the opportunity to display details about a signed signature line and display additional stored information such as a secure time-stamp.|
|[ShowSignatureSetup](../../Office.SignatureProvider.ShowSignatureSetup.md)|Provides a signature provider add-in the opportunity to display the **Signature Setup** dialog box to the user.|
|[ShowSigningCeremony](../../Office.SignatureProvider.ShowSigningCeremony.md)|Provides a signature provider add-in the opportunity to display the **Signature** dialog box to users, allowing them to specify their identity and then be authenticated.|
|[SignXmlDsig](../../Office.SignatureProvider.SignXmlDsig.md)|Used to sign the XMLDSIG template.|
|[VerifyXmlDsig](../../Office.SignatureProvider.VerifyXmlDsig.md)|Verifies a signature based on the signed state of the document and the legitimacy of the certificate used for signing.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]