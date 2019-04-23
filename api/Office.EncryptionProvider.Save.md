---
title: EncryptionProvider.Save method (Office)
keywords: vbaof11.chm327006
f1_keywords:
- vbaof11.chm327006
ms.prod: office
api_name:
- Office.EncryptionProvider.Save
ms.assetid: 7dfb6cea-f97b-51c3-e6bb-a773eec3fa73
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.Save method (Office)

Saves an encrypted document.


## Syntax

_expression_.**Save**(_SessionHandle_, _EncryptionData_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _EncryptionData_|Required|**IUnknown**|Contains the encryption information.|

## Return value

Long


## Remarks

When you save a file to the Office Open XML File format (which is the only format that supports custom file encryption), the provider is called by your COM add-in to encrypt the document. If you attempt to save to a format that does not support custom file encryption and you have the appropriate rights to do so, Microsoft Office will save the document without encryption. This allows documents to be exported to formats that do not support encryption or rights management.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]