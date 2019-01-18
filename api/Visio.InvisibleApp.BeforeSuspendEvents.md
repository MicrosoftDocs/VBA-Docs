---
title: InvisibleApp.BeforeSuspendEvents Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeSuspendEvents
ms.assetid: 6194a96a-d549-025b-fc97-7d79989447f7
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.BeforeSuspendEvents Event (Visio)

Occurs before firing of events is suspended.


## Syntax

Private Sub  _expression_ _'BeforeSuspendEvents'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ An expression that returns a [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events is going to be suspended.|

## Return value

nothing


## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


