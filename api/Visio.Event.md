---
title: Event object (Visio)
keywords: vis_sdr.chm10090
f1_keywords:
- vis_sdr.chm10090
ms.prod: visio
api_name:
- Visio.Event
ms.assetid: f11fffff-2218-8cc4-f223-31d956d1252d
ms.date: 06/19/2019
localization_priority: Normal
---


# Event object (Visio)

A member of the **[EventList](visio.eventlist.md)** collection of a source object such as a **Document**. An event encapsulates an event code.


## Remarks

An **Event** object can trigger two kinds of actions: it can run an add-on, or it can send a notification of the event to the calling program. To create an **Event** object, use the **Add** or **AddAdvise** method of an **EventList** object.

The default property of an **Event** object is **Event**.

The **Event** property of the **Event** object establishes the event that triggers the action, and its **Action** property indicates the action to be performed.

Use the **Persistable** property to find out if the event can be stored with a Microsoft Visio document, or the **Persistent** property to find out if the event is stored. 

Use the **Trigger** method to trigger an **Event** object's action without waiting for the event to occur. 

Use the **Enabled** property to temporarily disable an event.

## Methods

- [Delete](Visio.Event.Delete.md)
- [GetFilterActions](Visio.Event.GetFilterActions.md)
- [GetFilterCommands](Visio.Event.GetFilterCommands.md)
- [GetFilterObjects](Visio.Event.GetFilterObjects.md)
- [GetFilterSRC](Visio.Event.GetFilterSRC.md)
- [SetFilterActions](Visio.Event.SetFilterActions.md)
- [SetFilterCommands](Visio.Event.SetFilterCommands.md)
- [SetFilterObjects](Visio.Event.SetFilterObjects.md)
- [SetFilterSRC](Visio.Event.SetFilterSRC.md)
- [Trigger](Visio.Event.Trigger.md)

## Properties

- [Action](Visio.Event.Action.md)
- [Application](Visio.Event.Application.md)
- [Enabled](Visio.Event.Enabled.md)
- [Event](Visio.Event.Event.md)
- [EventList](Visio.Event.EventList.md)
- [ID](Visio.Event.ID.md)
- [Index](Visio.Event.Index.md)
- [ObjectType](Visio.Event.ObjectType.md)
- [Persistable](Visio.Event.Persistable.md)
- [Persistent](Visio.Event.Persistent.md)
- [Target](Visio.Event.Target.md)
- [TargetArgs](Visio.Event.TargetArgs.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]