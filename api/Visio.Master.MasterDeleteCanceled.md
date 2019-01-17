---
title: Master.MasterDeleteCanceled Event (Visio)
keywords: vis_sdr.chm10719355
f1_keywords:
- vis_sdr.chm10719355
ms.prod: visio
api_name:
- Visio.Master.MasterDeleteCanceled
ms.assetid: a682fab6-1fc9-65ba-83a1-408d048ee81e
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.MasterDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.


## Syntax

Private Sub  _expression_ _'MasterDeleteCanceled'(**_ByVal Master As [IVMASTER]_**)

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]