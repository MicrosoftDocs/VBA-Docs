---
title: Application.WindowActivated event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowActivated
ms.assetid: ef89f592-b457-b170-0e2e-84d9e1c572f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowActivated event (Visio)

Occurs after the active window changes in a Microsoft Visio instance.


## Syntax

_expression_.**WindowActivated** (_Window_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that was activated.|

## Remarks

The  **WindowActivated** event indicates that the active window has changed in a Visio instance. This event implies that the **ActiveDocument** and **ActivePage** properties of the **Application** object may also have changed; in contrast, any time the **ActiveDocument** or **ActivePage** property changes, a **WindowActivated** event is always generated.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this event maps to the following types:


-  **Microsoft.Office.Interop.Visio.EApplication_WindowActivatedEventHandler** (the **WindowActivated** delegate.)
    
-  **Microsoft.Office.Interop.Visio.EApplication_Event.WindowActivated** (the **WindowActivated** event.)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]