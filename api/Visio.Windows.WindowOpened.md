---
title: Windows.WindowOpened event (Visio)
keywords: vis_sdr.chm11719280
f1_keywords:
- vis_sdr.chm11719280
ms.prod: visio
api_name:
- Visio.Windows.WindowOpened
ms.assetid: 358b594b-0ca8-9d48-13a8-ae4dc527fdd6
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.WindowOpened event (Visio)

Occurs after a window is opened.


## Syntax

_expression_.**WindowOpened** (_Window_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that opened.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]