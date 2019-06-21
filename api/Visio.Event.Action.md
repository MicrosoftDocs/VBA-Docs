---
title: Event.Action property (Visio)
keywords: vis_sdr.chm12613010
f1_keywords:
- vis_sdr.chm12613010
ms.prod: visio
api_name:
- Visio.Event.Action
ms.assetid: dd776f54-051c-13c3-433e-299687203381
ms.date: 06/08/2017
localization_priority: Normal
---


# Event.Action property (Visio)

Gets or sets the action code of an **Event** object. Read/write.


## Syntax

_expression_.**Action**

_expression_ A variable that represents an **[Event](Visio.Event.md)** object.


## Return value

Integer


## Remarks

An **Event** object consists of an event-action pair; an event triggers an action. An action code is the numeric constant for the action that the event triggers.

Microsoft Visio supports the following action codes.

|Constant|Value|
|:-------|:---:|
| **visActCodeRunAddon**|1 |
| **visActCodeAdvise**|2 |


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]