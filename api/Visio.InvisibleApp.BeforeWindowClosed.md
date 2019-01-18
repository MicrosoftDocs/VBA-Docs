---
title: InvisibleApp.BeforeWindowClosed Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeWindowClosed
ms.assetid: b998ba51-0746-e7b1-263e-c93f98bd6f36
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.BeforeWindowClosed Event (Visio)

Occurs before a window is closed.


## Syntax

Private Sub  _expression_ _'BeforeWindowClosed'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to be closed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]