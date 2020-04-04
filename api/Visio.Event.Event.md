---
title: Event.Event property (Visio)
keywords: vis_sdr.chm12613470
f1_keywords:
- vis_sdr.chm12613470
ms.prod: visio
api_name:
- Visio.Event.Event
ms.assetid: 7b7783c3-2451-752e-6f40-ce25bd3fd696
ms.date: 06/08/2017
localization_priority: Normal
---


# Event.Event property (Visio)

Gets or sets the event code of an **Event** object—an event-action pair. When the event occurs, the action is performed. Read/write.


## Syntax

_expression_. `Event`

_expression_ A variable that represents an **[Event](Visio.Event.md)** object.


## Return value

Integer


## Remarks

If the action code of the  **Event** object is **visActCodeRunAddon**, the event also specifies the target of the action and the arguments to send to the target. This information is stored in the **Target** and **TargetArgs** properties, respectively.

If the action code of the  **Event** object is **visActCodeAdvise**, the event also specifies the object to receive event notifications (sometimes called the sink object) and arguments to send to the sink object along with the notification.

Event codes are declared by the Microsoft Visio type library in  **[VisEventCodes](Visio.viseventcodes.md)**. They are prefixed with " **visEvt** ". For a list of event codes, see [Event Codes](../visio/Concepts/event-codesvisio.md).

A program can use the  **Trigger** method to cause an **Event** object's action to be performed without waiting for the event to occur.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]