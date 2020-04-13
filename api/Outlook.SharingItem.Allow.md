---
title: SharingItem.Allow method (Outlook)
keywords: vbaol11.chm691
f1_keywords:
- vbaol11.chm691
ms.prod: outlook
api_name:
- Outlook.SharingItem.Allow
ms.assetid: 8f47e300-86d0-b90c-a41d-05bddec743f4
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Allow method (Outlook)

Allows a sharing request and sends a sharing response to the sender of the  **[SharingItem](Outlook.SharingItem.md)**.


## Syntax

_expression_. `Allow`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The **Allow** method can only be called on **SharingItem** objects with a **[Type](Outlook.SharingItem.Type.md)** property value of **olSharingMsgTypeRequest** or **olSharingMsgTypeInviteAndRequest**.

The **Type** property of the sharing response sent when this method is called is set to **olSharingMsgTypeResponseAllow**.


> [!NOTE] 
> Sharing is allowed immediately after this method is called, regardless of whether the sharing response was received.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]