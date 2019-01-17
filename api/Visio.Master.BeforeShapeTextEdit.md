---
title: Master.BeforeShapeTextEdit Event (Visio)
keywords: vis_sdr.chm10719380
f1_keywords:
- vis_sdr.chm10719380
ms.prod: visio
api_name:
- Visio.Master.BeforeShapeTextEdit
ms.assetid: 1d39001d-6efa-7d58-1eaa-f6c2531e2018
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.BeforeShapeTextEdit Event (Visio)

Occurs before a shape is opened for text editing in the user interface.


## Syntax

Private Sub  _expression_ _'BeforeShapeTextEdit'(**_ByVal Shape As [IVSHAPE]_**)

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is going to be opened for text editing.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


