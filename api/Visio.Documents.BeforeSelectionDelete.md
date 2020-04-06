---
title: Documents.BeforeSelectionDelete event (Visio)
keywords: vis_sdr.chm10619060
f1_keywords:
- vis_sdr.chm10619060
ms.prod: visio
api_name:
- Visio.Documents.BeforeSelectionDelete
ms.assetid: 0a0c764a-45bc-1ccb-a733-44c1933b95e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.BeforeSelectionDelete event (Visio)

Occurs before selected objects are deleted.


## Syntax

_expression_.**BeforeSelectionDelete** (_Selection_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selected objects that are going to be deleted.|

## Remarks

A  **Shape** object can serve as the source object for the **BeforeSelectionDelete** event if the shape's **Type** property is **visTypeGroup** (2) or **visTypePage** (1).

The  **BeforeSelectionDelete** event indicates that selected shapes are about to be deleted. This notification is sent whether or not any of the shapes are locked; however, locked shapes aren't deleted. To find out if a shape is locked against deletion, check the value of its LockDelete cell.

The  **BeforeSelectionDelete** and **BeforeShapeDelete** events are similar in that they both fire before shape(s) are deleted. They differ in how they behave when a single operation deletes several shapes. Suppose a **Cut** operation deletes three shapes. The **BeforeShapeDelete** event fires three times and acts on each of the three objects. The **BeforeSelectionDelete** event fires once, and it acts on a **Selection** object in which the three shapes that you want to delete are selected.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]