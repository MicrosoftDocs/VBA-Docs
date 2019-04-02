---
title: SharingItem.RemotePath property (Outlook)
keywords: vbaol11.chm696
f1_keywords:
- vbaol11.chm696
ms.prod: outlook
api_name:
- Outlook.SharingItem.RemotePath
ms.assetid: a0a40c81-2d57-1e6b-b565-70c765fcc829
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.RemotePath property (Outlook)

Returns a  **String** that represents the path of the sharing context for a **[SharingItem](Outlook.SharingItem.md)** object. Read-only.


## Syntax

_expression_. `RemotePath`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property contains the path of the sharing context, such as the URL of a WebCal calendar, contained within the  **SharingItem** object.

This property is set to an empty string if the  **[Type](Outlook.SharingItem.Type.md)** property of the **SharingItem** object is set to **olSharingMsgTypeRequest**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]