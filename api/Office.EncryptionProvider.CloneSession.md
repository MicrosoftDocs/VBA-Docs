---
title: EncryptionProvider.CloneSession method (Office)
keywords: vbaof11.chm327004
f1_keywords:
- vbaof11.chm327004
ms.prod: office
api_name:
- Office.EncryptionProvider.CloneSession
ms.assetid: d7548ad1-caec-27d8-db55-c4e6f747111e
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.CloneSession method (Office)

Creates a second, working copy of the **EncryptionProvider** object's encryption session for a file that is about to be saved.


## Syntax

_expression_.**CloneSession**(_SessionHandle_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the cloned session.|

## Return value

Long


## Remarks

The output of this method is another session handle with the same encryption settings. When an autoSave operation occurs, you are provided with this session handle.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]