---
title: Pages.ConnectionsAdded event (Visio)
keywords: vis_sdr.chm11019095
f1_keywords:
- vis_sdr.chm11019095
ms.prod: visio
api_name:
- Visio.Pages.ConnectionsAdded
ms.assetid: 7b2a471c-425f-0ab5-2cae-561dc67e343c
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.ConnectionsAdded event (Visio)

Occurs after connections have been established between shapes.


## Syntax

_expression_.**ConnectionsAdded** (_Connects_)

_expression_ A variable that represents a **[Pages](Visio.Pages.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Connects_|Required| **[IVCONNECTS]**|The connections that were established.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA **WithEvents** variables to sink the **ConnectionsDeleted** event.

For performance considerations, the  **Document** object's event set does not include the **ConnectionsAdded** event. To sink the **ConnectionsAdded** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]