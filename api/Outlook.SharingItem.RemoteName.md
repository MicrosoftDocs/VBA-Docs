---
title: SharingItem.RemoteName property (Outlook)
keywords: vbaol11.chm694
f1_keywords:
- vbaol11.chm694
ms.prod: outlook
api_name:
- Outlook.SharingItem.RemoteName
ms.assetid: 3c7fa798-cace-5748-3184-8055bf0f2958
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.RemoteName property (Outlook)

Returns a **String** that represents the name of the sharing context for a **[SharingItem](Outlook.SharingItem.md)** object. Read-only.


## Syntax

_expression_. `RemoteName`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property contains the name used for the sharing context, such as the name of the folder or item, contained within the  **SharingItem** object.

This property is set to an empty string if the  **[Type](Outlook.SharingItem.Type.md)** property of the **SharingItem** object is set to **olSharingMsgTypeRequest**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]