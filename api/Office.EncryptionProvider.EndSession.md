---
title: EncryptionProvider.EndSession method (Office)
keywords: vbaof11.chm327005
f1_keywords:
- vbaof11.chm327005
ms.prod: office
api_name:
- Office.EncryptionProvider.EndSession
ms.assetid: ce19f32e-a680-9d84-97d8-67d0f2d3b139
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.EndSession method (Office)

Ends the current encryption session.


## Syntax

_expression_.**EndSession**(_SessionHandle_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SessionHandle_|Required|**Long**|The ID of the current session.|

## Remarks

During a save operation, the **CloneSession** method is called by your COM add-in to create a second, working copy of the **EncryptionProvider** object's encryption session for the file that is about to be saved. Next, the **Save** method is called to get whatever custom information you would like to persist about your encryption settings. This information is available when this document is reopened later. 

The **EncryptStream** method is then called, which gives the provider the entire contents of the document. And finally, to complete the process, the **EndSession** method for the cloned session handle is called.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]