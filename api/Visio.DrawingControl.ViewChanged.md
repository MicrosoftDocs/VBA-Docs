---
title: DrawingControl.ViewChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.ViewChanged
ms.assetid: bab291c0-429a-bac5-339f-dcb71ce72199
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.ViewChanged Event (Visio)

Occurs when the zoom level or scroll position of a drawing window changes.


## Syntax

Private Sub  _expression_ _'ViewChanged'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents a [DrawingControl](./Visio.DrawingControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose zoom level or scroll position changed.|

## Remarks

This event fires whenever the zoom level or scroll position of a  **Window** object of the type **visDrawing** changes.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


