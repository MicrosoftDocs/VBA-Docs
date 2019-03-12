---
title: AcDataTransferType enumeration (Access)
keywords: vbaac10.chm10013
f1_keywords:
- vbaac10.chm10013
ms.prod: access
api_name:
- Access.AcDataTransferType
ms.assetid: cbd51e58-3873-ac1c-b494-55d43f1b2e25
ms.date: 06/08/2017
localization_priority: Normal
---


# AcDataTransferType enumeration (Access)

Specifies the type of transfer that you want to make with the **TransferDatabase** or **TransferSpreadsheet** method.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**acExport**|1|The data is exported.|
|**acImport**|0|(Default) The data is imported.|
|**acLink**|2|The database is linked to the specified data source.|

## Remarks

The **acLink** transfer type is not supported for Microsoft Access projects (.adp).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
