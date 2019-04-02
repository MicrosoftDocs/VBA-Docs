---
title: SharingItem.RemoteID property (Outlook)
keywords: vbaol11.chm695
f1_keywords:
- vbaol11.chm695
ms.prod: outlook
api_name:
- Outlook.SharingItem.RemoteID
ms.assetid: 07b0ba28-f560-7cee-bfc9-38fa073d8669
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.RemoteID property (Outlook)

Returns a  **String** that represents the unique identifier of the sharing context for a **[SharingItem](Outlook.SharingItem.md)** object. Read-only.


## Syntax

_expression_. `RemoteID`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property contains either a GUID or EntryID for the sharing context contained within the  **SharingItem** object.

This property is set to an empty string if the  **[Type](Outlook.SharingItem.Type.md)** property of the **SharingItem** object is set to **olSharingMsgTypeRequest**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]