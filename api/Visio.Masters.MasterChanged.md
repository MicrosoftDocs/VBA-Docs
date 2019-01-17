---
title: Masters.MasterChanged Event (Visio)
keywords: vis_sdr.chm10819175
f1_keywords:
- vis_sdr.chm10819175
ms.prod: visio
api_name:
- Visio.Masters.MasterChanged
ms.assetid: 824b7d27-b687-8a35-b97c-f4cf5e269065
ms.date: 06/08/2017
localization_priority: Normal
---


# Masters.MasterChanged Event (Visio)

Occurs after properties of a master are changed and propagated to its instances.


## Syntax

Private Sub  _expression_ _'MasterChanged'(**_ByVal Master As [IVMASTER]_**)

 _expression_ A variable that represents a [Masters](./Visio.Masters.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master whose properties changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


