---
title: InvisibleApp.BeforeStyleDelete event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeStyleDelete
ms.assetid: 0547897f-1ef9-27c4-1ea8-46e0e881ac91
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.BeforeStyleDelete event (Visio)

Occurs before a style is deleted.


## Syntax

_expression_.**BeforeStyleDelete** (_Style_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]