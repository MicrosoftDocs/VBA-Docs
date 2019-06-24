---
title: InvisibleApp.AppActivated event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AppActivated
ms.assetid: 8fb2624b-6755-c907-91b1-656f0031663f
ms.date: 06/24/2019
localization_priority: Normal
---


# InvisibleApp.AppActivated event (Visio)

Occurs after a Microsoft Visio instance becomes active.


## Syntax

_expression_.**AppActivated** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that becomes the active application.|

## Remarks

The **AppActivated** event indicates that an instance of Visio has become the active application on the Microsoft Windows desktop. The **AppActivated** event is different from the **AppObjectActivated** event, which occurs after an instance of Visio becomes active (the instance of Visio that is retrieved by the **GetObject** method in a Microsoft Visual Basic program).

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]