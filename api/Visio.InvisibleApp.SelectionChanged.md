---
title: InvisibleApp.SelectionChanged event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.SelectionChanged
ms.assetid: 6b621a79-9e12-ce5a-7cd0-9a2e6643d957
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.SelectionChanged event (Visio)

Occurs after a set of shapes selected in a window changes.


## Syntax

_expression_.**SelectionChanged** (_Window_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window in which the selection changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]