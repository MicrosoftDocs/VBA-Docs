---
title: Windows.BeforeWindowPageTurn event (Visio)
keywords: vis_sdr.chm11719080
f1_keywords:
- vis_sdr.chm11719080
ms.prod: visio
api_name:
- Visio.Windows.BeforeWindowPageTurn
ms.assetid: e74bbab7-af7b-19ef-af82-3f21b55a9292
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.BeforeWindowPageTurn event (Visio)

Occurs before a window is about to show a different page.


## Syntax

_expression_.**BeforeWindowPageTurn** (_Window_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to show a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]