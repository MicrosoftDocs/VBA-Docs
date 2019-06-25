---
title: Window.ViewChanged event (Visio)
keywords: vis_sdr.chm11619260
f1_keywords:
- vis_sdr.chm11619260
ms.prod: visio
api_name:
- Visio.Window.ViewChanged
ms.assetid: a65a8e2c-23d5-c582-cd42-4d6f4801d541
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.ViewChanged event (Visio)

Occurs when the zoom level or scroll position of a drawing window changes.


## Syntax

_expression_.**ViewChanged** (_Window_)

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose zoom level or scroll position changed.|

## Remarks

This event fires whenever the zoom level or scroll position of a  **Window** object of the type **visDrawing** changes.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]