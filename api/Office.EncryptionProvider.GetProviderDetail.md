---
title: EncryptionProvider.GetProviderDetail method (Office)
keywords: vbaof11.chm327001
f1_keywords:
- vbaof11.chm327001
ms.prod: office
api_name:
- Office.EncryptionProvider.GetProviderDetail
ms.assetid: d6bd91dc-ed35-bc75-9849-8caf989608d8
ms.date: 06/08/2017
---


# EncryptionProvider.GetProviderDetail method (Office)

Displays information about the encryption of the current document. 


## Syntax

_expression_. `GetProviderDetail`( `_encprovdet_` )

 _expression_ An expression that returns a [EncryptionProvider](Office.EncryptionProvider.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _encprovdet_|Required|**EncryptionProviderDetail**|Specifies the encryption information that you want.|

## Return value

Variant


## Remarks

This method allows you to query the  **EncryptionProvider** object for information such as what is the download URL for users without your custom COM add-in installed, what algorithm are you implementing, and what cipher mode you are using.


## See also


[EncryptionProvider Object](Office.EncryptionProvider.md)



[EncryptionProvider Object Members](./overview/Library-Reference/encryptionprovider-members-office.md)

