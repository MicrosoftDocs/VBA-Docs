---
title: AppointmentItem.Resources property (Outlook)
keywords: vbaol11.chm899
f1_keywords:
- vbaol11.chm899
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Resources
ms.assetid: 9b989d76-6897-cd2d-9156-fd7391dad8c1
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Resources property (Outlook)

Returns a semicolon-delimited  **String** of resource names for the meeting. Read/write.


## Syntax

_expression_. `Resources`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property contains the display names only. The  **[Recipients](Outlook.Recipients.md)** collection should be used to modify the resource recipients. Resources are added as **[BCC](Outlook.MailItem.BCC.md)** recipients to the collection.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]