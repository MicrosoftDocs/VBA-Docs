---
title: EncryptionProvider.NewSession method (Office)
keywords: vbaof11.chm327002
f1_keywords:
- vbaof11.chm327002
ms.prod: office
api_name:
- Office.EncryptionProvider.NewSession
ms.assetid: b90f842a-6eb3-3e95-7175-c3ca9c3ce138
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.NewSession method (Office)

Used by the **EncryptionProvider** object to create a new encryption session. This session is used by the provider to cache document-specific information about the encryption, users, and rights while the document is in memory.

## Syntax

_expression_.**NewSession**(_ParentWindow_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IUnknown**|Specifies the window that is called to display the encryption settings.|

## Return value

Long


## Remarks

This method is called by your COM add-in.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]