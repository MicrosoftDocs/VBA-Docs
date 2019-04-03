---
title: OlStorageIdentifierType enumeration (Outlook)
keywords: vbaol11.chm3111
f1_keywords:
- vbaol11.chm3111
ms.prod: outlook
api_name:
- Outlook.OlStorageIdentifierType
ms.assetid: 14283b38-6a0d-2954-bffe-87c36af27b2c
ms.date: 06/08/2017
localization_priority: Normal
---


# OlStorageIdentifierType enumeration (Outlook)

Specifies the type of identifier for a  **[StorageItem](Outlook.StorageItem.md)** object.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olIdentifyByEntryID**|1|Identifies a  **StorageItem** by **[EntryID](Outlook.StorageItem.EntryID.md)**.|
| **olIdentifyByMessageClass**|2|Identifies a  **StorageItem** by message class.|
| **olIdentifyBySubject**|0|Identifies a  **StorageItem** by **[Subject](Outlook.StorageItem.Subject.md)**.|

## Remarks

The message class of a [StorageItem object (Outlook)](Outlook.StorageItem.md) is not exposed as an explicit built-in property. You can access the message class property through the [PropertyAccessor object (Outlook)](Outlook.PropertyAccessor.md) that is provided by[StorageItem.PropertyAccessor property (Outlook)](Outlook.StorageItem.PropertyAccessor.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]