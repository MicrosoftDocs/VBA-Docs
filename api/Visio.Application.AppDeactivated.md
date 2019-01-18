---
title: Application.AppDeactivated Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.AppDeactivated
ms.assetid: 362bb2fb-91a2-01be-e686-3bf076388341
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AppDeactivated Event (Visio)

Occurs after a Microsoft Visio instance becomes inactive.


## Syntax

Private Sub  _expression_ _'AppDeactivated'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is no longer the active application.|

## Remarks

The  **AppDeactivated** event indicates that an instance of Visio is no longer the active application on the Microsoft Windows desktop. The **AppDeactivated** event is different from the **AppObjectDeactivated** event, which occurs after an instance of Visio ceases to be the active instance?the instance of Visio that is retrieved by the **GetObject** method in a Microsoft Visual Basic program.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]