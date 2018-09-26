---
title: InvisibleApp.AfterResumeEvents Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AfterResumeEvents
ms.assetid: 33117394-135e-0f44-79e8-d16531cd0ca5
ms.date: 06/08/2017
---


# InvisibleApp.AfterResumeEvents Event (Visio)

Occurs after firing of events is resumed.


## Syntax

 Private Sub _expression_ _'AfterResumeEvents'(**_ByVal app As [IVAPPLICATION]_** ,)

 _expression_ An expression that returns a [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events resumed.|

### Return Value

nothing


## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


