---
title: SignatureInfo.GetSignatureDetail method (Office)
keywords: vbaof11.chm286006
f1_keywords:
- vbaof11.chm286006
ms.prod: office
api_name:
- Office.SignatureInfo.GetSignatureDetail
ms.assetid: 77a5a835-cc8a-0341-8e5d-6ddb603f9517
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureInfo.GetSignatureDetail method (Office)

Displays a specified detail related to a signature.


## Syntax

_expression_.**GetSignatureDetail**(_sigdet_)

_expression_ An expression that returns a **[SignatureInfo](Office.SignatureInfo.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _sigdet_|Required|**SignatureDetail**|An enumerated value specifying which signature detail to display.|

## Return value

Variant


## Example

The following example gets information on the suggested signer of the document.


```vb
Sub GetSigDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetSignatureDetail(sigdetDelSuggSigner) 
 
End Sub
```


## See also

- [SignatureInfo object members](overview/Library-Reference/signatureinfo-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]