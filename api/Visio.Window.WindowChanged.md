---
title: Window.WindowChanged Event (Visio)
keywords: vis_sdr.chm11619275
f1_keywords:
- vis_sdr.chm11619275
ms.prod: visio
api_name:
- Visio.Window.WindowChanged
ms.assetid: ee7e4871-26ca-ea4e-1c7b-2e597d92e143
ms.date: 06/08/2017
---


# Window.WindowChanged Event (Visio)

Occurs when the size or position of a window changes.


## Syntax

Private Sub  _expression_ _'WindowChanged'(**_ByVal Window As [IVWINDOW]_**)

 _expression_ A variable that represents a [Window](./Visio.Window.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window whose size or position has changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


