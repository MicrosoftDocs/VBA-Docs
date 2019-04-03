---
title: OlkTimeZoneControl.Exit event (Outlook)
keywords: vbaol11.chm1000522
f1_keywords:
- vbaol11.chm1000522
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.Exit
ms.assetid: da5616c5-97da-6049-4115-5a41d4e28c7b
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeZoneControl.Exit event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. `Exit`( `_Cancel_` )

_expression_ A variable that represents an [OlkTimeZoneControl](Outlook.OlkTimeZoneControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains on this control|

## See also


[OlkTimeZoneControl Object](Outlook.OlkTimeZoneControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]