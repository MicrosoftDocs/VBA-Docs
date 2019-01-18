---
title: Application.BeforeMasterDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeMasterDelete
ms.assetid: 65e3bbed-46f4-25c1-1e3f-af61ef61cce9
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BeforeMasterDelete Event (Visio)

Occurs before a master is deleted from a document.


## Syntax

Private Sub  _expression_ _'BeforeMasterDelete'(**_ByVal Master As [IVMASTER]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]