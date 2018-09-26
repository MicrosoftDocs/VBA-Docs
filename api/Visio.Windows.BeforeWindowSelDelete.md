---
title: Windows.BeforeWindowSelDelete Event (Visio)
keywords: vis_sdr.chm11719085
f1_keywords:
- vis_sdr.chm11719085
ms.prod: visio
api_name:
- Visio.Windows.BeforeWindowSelDelete
ms.assetid: db81302b-bfc9-672d-9a73-45fe34f89136
ms.date: 06/08/2017
---


# Windows.BeforeWindowSelDelete Event (Visio)

Occurs before the shapes in the selection of a window are deleted.


## Syntax

Private Sub  _expression_ _'BeforeWindowSelDelete'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents a [Windows](./Visio.Windows.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that contains the selection that is going to be deleted.|

## Remarks

The  **BeforeWindowSelDelete** event fires if user interactions cause shapes in a window to be deleted. It doesn't fire if a program deletes shapes in a window by using the **Cut** method, for example.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


