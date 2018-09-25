---
title: DrawingControl.ConvertToGroupCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.ConvertToGroupCanceled
ms.assetid: de4c6838-62dd-c983-3677-a8598c09edeb
ms.date: 06/08/2017
---


# DrawingControl.ConvertToGroupCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.


## Syntax

Private Sub  _expression_ _'ConvertToGroupCanceled'(**_ByVal selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [DrawingControl](./Visio.DrawingControl.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be grouped.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


