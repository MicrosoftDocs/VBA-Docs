---
title: OlAppointmentCopyOptions enumeration (Outlook)
keywords: vbaol11.chm3513
f1_keywords:
- vbaol11.chm3513
ms.prod: outlook
api_name:
- Outlook.OlAppointmentCopyOptions
ms.assetid: b2ea721d-f800-6102-c893-28f265e70b88
ms.date: 06/08/2017
localization_priority: Normal
---


# OlAppointmentCopyOptions enumeration (Outlook)

Specifies what actions to take when copying an  **[AppointmentItem](Outlook.AppointmentItem.md)** object to a folder.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olCopyAsAccept**|2|Creates an appointment in the destination folder and accepts the meeting request automatically.|
| **olCreateAppointment**|1|Creates an appointment in the destination folder without defaulting to a response or prompting for a response.|
| **olPromptUser**|0|Copies the appointment to the destination folder and prompts the user to accept the request before completing the copy operation.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]