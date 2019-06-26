---
title: InvisibleApp.NoEventsPending event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.NoEventsPending
ms.assetid: 65947eae-69de-3220-e4e5-5edf5b6ad242
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.NoEventsPending event (Visio)

Occurs after the Microsoft Visio instance flushes its event queue.


## Syntax

_expression_.**NoEventsPending** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that flushed its event queue.|

## Remarks

Visio maintains a queue of events and fires them at discrete moments. Immediately after Visio fires the last event in its event queue, it fires a **NoEventsPending** event.

A client program can use the **NoEventsPending** event as a signal that Visio has completed a burst of activity. For example, a client program may want to react to changes in a shape's geometry. A single user action performed on the shape can generate several **CellChanged** events. The client program could record selected information for each **CellChanged** event and perform its processing after it receives the **NoEventsPending** event.

Visio fires the **NoEventsPending** event only if at least one of the events in the queue is being listened to. If no program is listening for any of the queued events, the **NoEventsPending** event does not fire. If your program is only listening to the **NoEventsPending** event, it does not fire unless another program is listening for some of the queued events.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]