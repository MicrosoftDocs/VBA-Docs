---
title: OlkSenderPhoto.Change event (Outlook)
keywords: vbaol11.chm1000492
f1_keywords:
- vbaol11.chm1000492
ms.prod: outlook
api_name:
- Outlook.OlkSenderPhoto.Change
ms.assetid: a4d58172-a16f-6084-9230-af2c3cefa552
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkSenderPhoto.Change event (Outlook)

Occurs when the sender's contact picture has changed. 


## Syntax

_expression_. `Change`

_expression_ A variable that represents an [OlkSenderPhoto](Outlook.OlkSenderPhoto.md) object.


## Remarks

The change of the sender's contact picture usually means that the  **[PreferredWidth](Outlook.OlkSenderPhoto.PreferredWidth.md)** and **[PreferredHeight](Outlook.OlkSenderPhoto.PreferredHeight.md)** properties have changed as well. Therefore, this event can be used as an indication of the necessity to resize the control.


## See also


[OlkSenderPhoto Object](Outlook.OlkSenderPhoto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]