---
title: Application.WindowCloseCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowCloseCanceled
ms.assetid: 1273b75d-0543-69aa-aab3-47281295ee6b
ms.date: 06/08/2017
---


# Application.WindowCloseCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelWindowClose** event.


## Syntax

Private Sub  _expression_ _'WindowCloseCanceled'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that was going to be closed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


