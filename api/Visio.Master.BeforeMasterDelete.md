---
title: Master.BeforeMasterDelete event (Visio)
keywords: vis_sdr.chm10719040
f1_keywords:
- vis_sdr.chm10719040
ms.prod: visio
api_name:
- Visio.Master.BeforeMasterDelete
ms.assetid: 46b455db-9165-0ed4-ebf3-15e1794313be
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.BeforeMasterDelete event (Visio)

Occurs before a master is deleted from a document.


## Syntax

_expression_.**BeforeMasterDelete** (_Master_)

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]