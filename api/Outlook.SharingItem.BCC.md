---
title: SharingItem.BCC property (Outlook)
keywords: vbaol11.chm633
f1_keywords:
- vbaol11.chm633
ms.prod: outlook
api_name:
- Outlook.SharingItem.BCC
ms.assetid: e13c7fab-5ce6-289a-35d0-ffea5d0bd09e
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.BCC property (Outlook)

Returns a  **String** representing the display list of blind carbon copy (BCC) names for a **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `BCC`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property contains only the display names, delimited with semicolon (;) characters. The  **[Recipients](Outlook.Recipients.md)** collection should be used to modify the BCC recipients.


> [!NOTE] 
> If the  **SharingItem** uses an Exchange sharing context, then setting this property to any value other than **Nothing** prevents the item from being sent and causes the **[Send](Outlook.SharingItem.Send(method).md)** method to raise an error.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]