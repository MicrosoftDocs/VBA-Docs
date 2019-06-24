---
title: Page.GroupCanceled event (Visio)
keywords: vis_sdr.chm10962005
f1_keywords:
- vis_sdr.chm10962005
ms.prod: visio
api_name:
- Visio.Page.GroupCanceled
ms.assetid: ae7bd6b5-8975-26a2-86af-ff12eaef5ebb
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.GroupCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.


## Syntax

_expression_.**GroupCanceled** (_Selection_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be grouped.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]