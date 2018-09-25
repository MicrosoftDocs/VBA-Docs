---
title: Window.BeforeWindowClosed Event (Visio)
keywords: vis_sdr.chm11619075
f1_keywords:
- vis_sdr.chm11619075
ms.prod: visio
api_name:
- Visio.Window.BeforeWindowClosed
ms.assetid: 4543e237-6b2c-d02c-66df-9f90b0266e4b
ms.date: 06/08/2017
---


# Window.BeforeWindowClosed Event (Visio)

Occurs before a window is closed.


## Syntax

Private Sub  _expression_ _'BeforeWindowClosed'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents a [Window](./Visio.Window.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to be closed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


