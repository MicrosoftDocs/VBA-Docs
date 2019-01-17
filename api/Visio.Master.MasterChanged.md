---
title: Master.MasterChanged Event (Visio)
keywords: vis_sdr.chm10719175
f1_keywords:
- vis_sdr.chm10719175
ms.prod: visio
api_name:
- Visio.Master.MasterChanged
ms.assetid: 922120cc-56e0-143b-7a8b-754bc368af47
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.MasterChanged Event (Visio)

Occurs after properties of a master are changed and propagated to its instances.


## Syntax

Private Sub  _expression_ _'MasterChanged'(**_ByVal Master As [IVMASTER]_**)

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master whose properties changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


