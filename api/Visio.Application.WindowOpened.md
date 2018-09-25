---
title: Application.WindowOpened Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowOpened
ms.assetid: a75a50b5-9784-e191-991a-ca9b41994ff9
ms.date: 06/08/2017
---


# Application.WindowOpened Event (Visio)

Occurs after a window is opened.


## Syntax

Private Sub  _expression_ _'WindowOpened'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that opened.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


