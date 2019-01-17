---
title: InvisibleApp.BeforeWindowPageTurn Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeWindowPageTurn
ms.assetid: 7bcdf0f3-2825-9205-06fb-252170018073
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.BeforeWindowPageTurn Event (Visio)

Occurs before a window is about to show a different page.


## Syntax

Private Sub  _expression_ _'BeforeWindowPageTurn'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to show a different page.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


