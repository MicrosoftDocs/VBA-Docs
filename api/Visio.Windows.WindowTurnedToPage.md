---
title: Windows.WindowTurnedToPage event (Visio)
keywords: vis_sdr.chm11719285
f1_keywords:
- vis_sdr.chm11719285
ms.prod: visio
api_name:
- Visio.Windows.WindowTurnedToPage
ms.assetid: cf0f0170-41ab-92a7-1fe3-e0617af48b0d
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.WindowTurnedToPage event (Visio)

Occurs after a window shows a different page.


## Syntax

_expression_.**WindowTurnedToPage** (_Window_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that shows a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]