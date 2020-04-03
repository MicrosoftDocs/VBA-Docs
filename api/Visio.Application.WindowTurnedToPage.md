---
title: Application.WindowTurnedToPage event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowTurnedToPage
ms.assetid: f747ed48-6da1-fd7f-4cdd-e9f46f02b1d0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowTurnedToPage event (Visio)

Occurs after a window shows a different page.


## Syntax

_expression_.**WindowTurnedToPage** (_Window_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


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

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this event maps to the following types:


-  **Microsoft.Office.Interop.Visio.EApplication_WindowTurnedToPageEventHandler** (the **WindowTurnedToPage** delegate.)
    
-  **Microsoft.Office.Interop.Visio.EApplication_Event.WindowTurnedToPage** (the **WindowTurnedToPage** event.)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]