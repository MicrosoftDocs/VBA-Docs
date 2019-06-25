---
title: Windows.WindowChanged event (Visio)
keywords: vis_sdr.chm11719275
f1_keywords:
- vis_sdr.chm11719275
ms.prod: visio
api_name:
- Visio.Windows.WindowChanged
ms.assetid: 02893ec6-2aca-cb70-919d-6ea2d37bb915
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.WindowChanged event (Visio)

Occurs when the size or position of a window changes.


## Syntax

_expression_.**WindowChanged** (_Window_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose size or position has changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]