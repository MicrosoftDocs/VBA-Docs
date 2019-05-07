---
title: Pages.UngroupCanceled event (Visio)
keywords: vis_sdr.chm11019375
f1_keywords:
- vis_sdr.chm11019375
ms.prod: visio
api_name:
- Visio.Pages.UngroupCanceled
ms.assetid: 9ee7c970-7cb4-3683-b71c-1c828bbd4ec4
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.UngroupCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.


## Syntax

Private Sub  _expression_ _'UngroupCanceled'(**_ByVal Selection As [IVSELECTION]_**)

_expression_ A variable that represents a [Pages](./Visio.Pages.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be ungrouped.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]