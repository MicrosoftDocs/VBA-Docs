---
title: Application.WindowTurnedToPage Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.WindowTurnedToPage
ms.assetid: f747ed48-6da1-fd7f-4cdd-e9f46f02b1d0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowTurnedToPage Event (Visio)

Occurs after a window shows a different page.


## Syntax

Private Sub  _expression_ _'WindowTurnedToPage'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that shows a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this event maps to the following types:


-  **Microsoft.Office.Interop.Visio.EApplication_WindowTurnedToPageEventHandler** (the **WindowTurnedToPage** delegate.)
    
-  **Microsoft.Office.Interop.Visio.EApplication_Event.WindowTurnedToPage** (the **WindowTurnedToPage** event.)
    

