---
title: InvisibleApp.BeforeSelectionDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeSelectionDelete
ms.assetid: a5407295-e103-dc71-6655-df708aa905a3
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.BeforeSelectionDelete Event (Visio)

Occurs before selected objects are deleted.


## Syntax

Private Sub  _expression_ _'BeforeSelectionDelete'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selected objects that are going to be deleted.|

## Remarks

A  **Shape** object can serve as the source object for the **BeforeSelectionDelete** event if the shape's **Type** property is **visTypeGroup** (2) or **visTypePage** (1).

The  **BeforeSelectionDelete** event indicates that selected shapes are about to be deleted. This notification is sent whether or not any of the shapes are locked; however, locked shapes aren't deleted. To find out if a shape is locked against deletion, check the value of its LockDelete cell.

The  **BeforeSelectionDelete** and **BeforeShapeDelete** events are similar in that they both fire before shape(s) are deleted. They differ in how they behave when a single operation deletes several shapes. Suppose a **Cut** operation deletes three shapes. The **BeforeShapeDelete** event fires three times and acts on each of the three objects. The **BeforeSelectionDelete** event fires once, and it acts on a **Selection** object in which the three shapes that you want to delete are selected.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


