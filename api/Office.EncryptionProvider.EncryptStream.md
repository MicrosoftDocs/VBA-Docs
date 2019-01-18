---
title: EncryptionProvider.EncryptStream method (Office)
keywords: vbaof11.chm327007
f1_keywords:
- vbaof11.chm327007
ms.prod: office
api_name:
- Office.EncryptionProvider.EncryptStream
ms.assetid: 58a379f4-fb74-4a2c-b0ed-ce3e3151c292
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.EncryptStream method (Office)

Encrypts and returns a stream of data for a document.


## Syntax

_expression_.**EncryptStream**(_SessionHandle_, _StreamName_, _UnencryptedStream_, _EncryptedStream_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|
| _StreamName_|Required|**String**|The name of the encrypted stream of document data.|
| _UnencryptedStream_|Required|**IUnknown**|The data stream before encryption.|
| _EncryptedStream_|Required|**IUnknown**|The data stream information after it has been encrypted.|

## Remarks

This method is typically called by your COM add-in during a save operation.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]