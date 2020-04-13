---
title: SharingItem.CC property (Outlook)
keywords: vbaol11.chm634
f1_keywords:
- vbaol11.chm634
ms.prod: outlook
api_name:
- Outlook.SharingItem.CC
ms.assetid: ac3e12ea-6e3d-71c8-ecb4-c7d54d669cee
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.CC property (Outlook)

Returns a **String** representing the display list of carbon copy (CC) names for a **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `CC`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property contains only the display names, delimited by semicolon (;) characters. The **[Recipients](Outlook.Recipients.md)** collection should be used to modify the CC recipients.


> [!NOTE] 
> If the  **SharingItem** uses an Exchange sharing context, then setting this property to any value other than **Nothing** prevents the item from being sent and causes the **[Send](Outlook.SharingItem.Send(method).md)** method to raise an error.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]