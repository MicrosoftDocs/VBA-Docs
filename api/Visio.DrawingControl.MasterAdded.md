---
title: DrawingControl.MasterAdded event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.MasterAdded
ms.assetid: 9c462361-fa77-2916-b2f1-e7b064754bc1
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.MasterAdded event (Visio)

Occurs after a new master is added to a document.


## Syntax

_expression_.**MasterAdded** (_Master_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was added to the document.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]