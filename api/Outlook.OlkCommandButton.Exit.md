---
title: OlkCommandButton.Exit event (Outlook)
keywords: vbaol11.chm1000126
f1_keywords:
- vbaol11.chm1000126
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.Exit
ms.assetid: be3f7740-8682-ecc5-3927-dd700f26b49c
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCommandButton.Exit event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. `Exit`( `_Cancel_` )

_expression_ A variable that represents an [OlkCommandButton](Outlook.OlkCommandButton.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains in this control.|

## See also


[OlkCommandButton Object](Outlook.OlkCommandButton.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]