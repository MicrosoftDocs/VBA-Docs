---
title: InvisibleApp.ConnectionsDeleted Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ConnectionsDeleted
ms.assetid: 88505099-3b7d-bf02-cc3d-d56bc436e63f
ms.date: 06/08/2017
---


# InvisibleApp.ConnectionsDeleted Event (Visio)

Occurs after connections between shapes have been removed.


## Syntax

Private Sub  _expression_ _'ConnectionsDeleted'(**_ByVal Connects As [IVCONNECTS]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Connects_|Required| **[IVCONNECTS]**|The connections that have been removed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




 **Note**   You can use VBA **WithEvents** variables to sink the **ConnectionsDeleted** event.

For performance considerations, the  **Document** object's event set does not include the **ConnectionsDeleted** event. To sink the **ConnectionsDeleted** event from a **Document** object (and the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.


