---
title: Application.WindowChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowChanged
ms.assetid: 29bb6ea8-4558-38c4-941f-839cd119abba
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowChanged Event (Visio)

Occurs when the size or position of a window changes.


## Syntax

Private Sub  _expression_ _'WindowChanged'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose size or position has changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


