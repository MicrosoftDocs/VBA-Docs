---
title: Masters.GroupCanceled Event (Visio)
keywords: vis_sdr.chm10862005
f1_keywords:
- vis_sdr.chm10862005
ms.prod: visio
api_name:
- Visio.Masters.GroupCanceled
ms.assetid: dbdecd35-1996-465d-afd3-a82e6bb14f7b
ms.date: 06/08/2017
---


# Masters.GroupCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.


## Syntax

Private Sub  _expression_ _'GroupCanceled'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Masters](./Visio.Masters.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be grouped.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


