---
title: SignatureProvider.GetProviderDetail method (Office)
keywords: vbaof11.chm287008
f1_keywords:
- vbaof11.chm287008
ms.prod: office
api_name:
- Office.SignatureProvider.GetProviderDetail
ms.assetid: a8cc567e-be67-3a5e-d719-40da6d294fb4
ms.date: 06/08/2017
localization_priority: Normal
---


# SignatureProvider.GetProviderDetail method (Office)

Queries the signature provider add-in for various details. 


## Syntax

_expression_. `GetProviderDetail`( `_sigprovdet_` )

 _expression_ An expression that returns a [SignatureProvider](Office.SignatureProvider.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _sigprovdet_|Required|**SignatureProviderDetail**|Contains an enumerated value representing the type of information to query the add-in for.|

## Return value

Variant


## Remarks

The  **SignatureProvider** object is used exclusively in custom signature provider add-ins. This method is used to query the add-in for three pieces of information:


- What hash algorithm does the add-in support?
    
- Is the add-in only a user interface (UI) or does it support hashing and verification? If  **TRUE** is returned, Microsoft Office does not call the add-in to hash or verify, only to display the UI.
    
- What URL should the add-in provide for users if they are missing the signature add-in?
    



## Example

The following example, written in C#, shows the implementation of the  **GetProviderDetail** method in a custom signature provider project.


```vb
 public object GetProviderDetail(SignatureProviderDetail sigProvDetail) 
 { 
 switch (sigProvDetail) 
 { 
 case Microsoft.Office.Core.SignatureProviderDetail.sigprovdetHashAlgorithm: 
 return this.HashAlgorithmIdentifier; 
 
 case Microsoft.Office.Core.SignatureProviderDetail.sigprovdetUIOnly: 
 return false; 
 
 case Microsoft.Office.Core.SignatureProviderDetail.sigprovdetUrl: 
 return this.ProviderUrl; 
 
 default: 
 return null; 
 } 
 } 

```

> [!NOTE] 
> Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


[SignatureProvider Object](Office.SignatureProvider.md)



[SignatureProvider Object Members](./overview/Library-Reference/signatureprovider-members-office.md)

