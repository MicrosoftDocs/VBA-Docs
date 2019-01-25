---
title: SignatureSet.AddNonVisibleSignature method (Office)
keywords: vbaof11.chm247006
f1_keywords:
- vbaof11.chm247006
ms.prod: office
api_name:
- Office.SignatureSet.AddNonVisibleSignature
ms.assetid: f8d3a749-9507-628f-2192-552bd4cbb00c
ms.date: 01/24/2019
localization_priority: Normal
---


# SignatureSet.AddNonVisibleSignature method (Office)

Creates a signature packet when digitally signing a document.


## Syntax

_expression_.**AddNonVisibleSignature**(_varSigProv_)

_expression_ An expression that returns a **[SignatureSet](Office.SignatureSet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _varSigProv_|Optional|**Variant**|Represents the ID of the signature provider.|

## Return value

Signature


## Remarks

To provide an entry point to trigger this method, you need to create a user interface with the signature provider add-in. This entry point is typically provided to the user as a menu option.


## Example

The following function uses the signature provider ID argument to create a signature packet when digitally signing a document.


```vb
Function CreateSignature(ByVal varSigProviderID As Variant) As Signature 
Dim objSignatureSet As SignatureSet 
Dim objSignature As Signature 
 
objSignature = objSignatureSet.AddNonVisibleSignature(varSigProviderID) 
CreateSignature = objSignature 
 
End Function
```


## See also

- [SignatureSet object members](overview/Library-Reference/signatureset-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]