---
title: Signature.Sign method (Office)
keywords: vbaof11.chm248012
f1_keywords:
- vbaof11.chm248012
ms.prod: office
api_name:
- Office.Signature.Sign
ms.assetid: 37ba202a-da6d-9978-c8af-986a8218e004
ms.date: 01/24/2019
localization_priority: Normal
---


# Signature.Sign method (Office)

Creates a signature packet.


## Syntax

_expression_.**Sign** (_varSigImg_, _varDelSuggSigner_, _varDelSuggSignerLine2_, _varDelSuggSignerEmail_)

 _expression_ An expression that returns a **[Signature](Office.Signature.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _varSigImg_|Optional|**Variant**|The signature line graphic image.|
| _varDelSuggSigner_|Optional|**Variant**|The suggested signer.|
| _varDelSuggSignerLine2_|Optional|**Variant**|The additional signature line.|
| _varDelSuggSignerEmail_|Optional|**Variant**|The email address of the suggested signer.|

## Remarks

When the **Sign** method is called, Microsoft Office creates a manifest and calls the signature provider to create a hash for each stream in the document. Office then bundles up the results into an unsigned XMLDSIG template and calls to the provider to modify the XMLDSIG (if necessary) and then sign it. The resulting signed signature is then handed back to Office to be stored.


## Example

In the following example, the variables for the signature image, signer, signer's title, and signer's email address are set, and then the **Sign** method is called to create and sign a signature packet.


```vb
Set objSignature = New Signature 
varSigline = CType(AxHost2.GetIPictureDispFromPicture(img),IPictureDisp) 
varSuggestedSigner = "Nancy Davolio" 
varSignatureTitle = "Sales Representative" 
varSignerEmail = "ndavolio@northwindtraders.com" 
objSignature.Sign(varSigline, varSuggestedSigner, varSignatureTitle, varSignerEmail)
```


## See also

- [Signature object members](overview/Library-Reference/signature-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]