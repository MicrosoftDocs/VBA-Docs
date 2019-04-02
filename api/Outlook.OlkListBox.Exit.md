---
title: OlkListBox.Exit event (Outlook)
keywords: vbaol11.chm1000286
f1_keywords:
- vbaol11.chm1000286
ms.prod: outlook
api_name:
- Outlook.OlkListBox.Exit
ms.assetid: 729d454a-4f52-c0c2-4125-7cbf8ea2d660
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkListBox.Exit event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. `Exit`( `_Cancel_` )

_expression_ A variable that represents an [OlkListBox](Outlook.OlkListBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains in this control.|

## See also


[OlkListBox Object](Outlook.OlkListBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]