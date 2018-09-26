---
title: InvisibleApp.WindowTurnedToPage Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.WindowTurnedToPage
ms.assetid: a31992e8-7b3e-2986-a9e8-01cae1ae1fa5
ms.date: 06/08/2017
---


# InvisibleApp.WindowTurnedToPage Event (Visio)

Occurs after a window shows a different page.


## Syntax

Private Sub  _expression_ _'WindowTurnedToPage'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that shows a different page.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


