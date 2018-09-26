---
title: InvisibleApp.AppActivated Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.AppActivated
ms.assetid: 8fb2624b-6755-c907-91b1-656f0031663f
ms.date: 06/08/2017
---


# InvisibleApp.AppActivated Event (Visio)

Occurs after a Microsoft Visio instance becomes active.


## Syntax

Private Sub  _expression_ _'AppActivated'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that becomes the active application.|

## Remarks

The  **AppActivated** event indicates that an instance of Visio has become the active application on the Microsoft Windows desktop. The **AppActivated** event is different from the **AppObjectActivated** event, which occurs after an instance of Visio becomes active?the instance of Visio that is retrieved by the **GetObject** method in a Microsoft Visual Basic program.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


