---
title: Application.BeforeWindowSelDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeWindowSelDelete
ms.assetid: 36ff6935-23a8-b155-e5d1-58ae90b10cb6
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BeforeWindowSelDelete Event (Visio)

Occurs before the shapes in the selection of a window are deleted.


## Syntax

Private Sub  _expression_ _'BeforeWindowSelDelete'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that contains the selection that is going to be deleted.|

## Remarks

The  **BeforeWindowSelDelete** event fires if user interactions cause shapes in a window to be deleted. It doesn't fire if a program deletes shapes in a window by using the **Cut** method, for example.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]