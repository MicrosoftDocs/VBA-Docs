---
title: Application.MasterDeleteCanceled event (Visio)
ms.prod: visio
api_name:
- Visio.Application.MasterDeleteCanceled
ms.assetid: 8dabb35b-8959-ef83-90fd-3287265f60a5
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.MasterDeleteCanceled event (Visio)

Occurs after an event handler has returned **True** (cancel) to a **QueryCancelMasterDelete** event.


## Syntax

_expression_.**MasterDeleteCanceled** (_Master_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]