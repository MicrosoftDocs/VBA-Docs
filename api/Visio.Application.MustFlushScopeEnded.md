---
title: Application.MustFlushScopeEnded Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.MustFlushScopeEnded
ms.assetid: ba9ae16a-9cc6-79d6-d838-e5927937c142
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MustFlushScopeEnded Event (Visio)

Occurs after the Microsoft Visio instance is forced to flush its event queue.


## Syntax

Private Sub  _expression_ _'MustFlushScopeEnded'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is forced to flush its event queue.|

## Remarks

This event, along with the  **MustFlushScopeBeginning** event, can be used to identify whether an event is being fired because Visio is forced to flush its event queue.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]