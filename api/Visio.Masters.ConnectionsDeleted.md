---
title: Masters.ConnectionsDeleted event (Visio)
keywords: vis_sdr.chm10819100
f1_keywords:
- vis_sdr.chm10819100
ms.prod: visio
api_name:
- Visio.Masters.ConnectionsDeleted
ms.assetid: bf2ed2be-276a-04d8-cd98-70929cfd31f6
ms.date: 06/08/2017
localization_priority: Normal
---


# Masters.ConnectionsDeleted event (Visio)

Occurs after connections between shapes have been removed.


## Syntax

_expression_.**ConnectionsDeleted** (_Connects_)

_expression_ A variable that represents a **[Masters](Visio.Masters.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Connects_|Required| **[IVCONNECTS]**|The connections that have been removed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA **WithEvents** variables to sink the **ConnectionsDeleted** event.

For performance considerations, the  **Document** object's event set does not include the **ConnectionsDeleted** event. To sink the **ConnectionsDeleted** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]