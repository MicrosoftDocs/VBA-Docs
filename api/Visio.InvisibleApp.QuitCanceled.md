---
title: InvisibleApp.QuitCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QuitCanceled
ms.assetid: 48e46a44-581f-cd79-dbeb-6ee70c6b391b
ms.date: 06/08/2017
---


# InvisibleApp.QuitCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelQuit** event.


## Syntax

Private Sub  _expression_ _'QuitCanceled'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that was going to be terminated.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


