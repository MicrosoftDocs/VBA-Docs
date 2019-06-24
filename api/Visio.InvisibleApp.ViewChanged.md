---
title: InvisibleApp.ViewChanged event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ViewChanged
ms.assetid: 6c69423e-50fa-af13-b198-bb628df17440
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.ViewChanged event (Visio)

Occurs when the zoom level or scroll position of a drawing window changes.


## Syntax

_expression_.**ViewChanged** (_Window_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


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