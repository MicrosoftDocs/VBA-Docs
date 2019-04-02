---
title: OlkContactPhoto.Exit event (Outlook)
keywords: vbaol11.chm1000317
f1_keywords:
- vbaol11.chm1000317
ms.prod: outlook
api_name:
- Outlook.OlkContactPhoto.Exit
ms.assetid: 8bc0e21f-7376-3bc7-5006-a00031686229
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkContactPhoto.Exit event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. `Exit`( `_Cancel_` )

_expression_ A variable that represents an [OlkContactPhoto](Outlook.OlkContactPhoto.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains in this control.|

## Remarks

This event occurs only when the control is displaying the contact picture button and no contact picture has been defined.


## See also


[OlkContactPhoto Object](Outlook.OlkContactPhoto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]