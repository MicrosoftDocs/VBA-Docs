---
title: Inspector.BeforeMinimize event (Outlook)
keywords: vbaol11.chm469
f1_keywords:
- vbaol11.chm469
ms.prod: outlook
api_name:
- Outlook.Inspector.BeforeMinimize
ms.assetid: a2a6ce7e-5980-2914-6785-be87d9b163c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.BeforeMinimize event (Outlook)

Occurs when the active inspector is minimized by the user.


## Syntax

_expression_. `BeforeMinimize`( `_Cancel_` )

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the explorer or inspector is not minimized.|

## Remarks

This event can be cancelled after it has started.


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]