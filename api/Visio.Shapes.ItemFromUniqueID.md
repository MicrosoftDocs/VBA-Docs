---
title: Shapes.ItemFromUniqueID property (Visio)
keywords: vis_sdr.chm11362485
f1_keywords:
- vis_sdr.chm11362485
ms.prod: visio
api_name:
- Visio.Shapes.ItemFromUniqueID
ms.assetid: 94175764-d65d-9511-4073-864ff89f573c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.ItemFromUniqueID property (Visio)

Returns the  **[Shape](Visio.Shape.md)** object that matches the specified **[UniqueID](Visio.Shape.UniqueID.md)** property value. Read-only.


## Syntax

_expression_. `ItemFromUniqueID`( `_UniqueID_` )

_expression_ A variable that represents a **[Shapes](Visio.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UniqueID_|Required| **String**|The unique ID of a  **Shape** object.|

## Return value

 **Shape**


## Remarks

Microsoft Visio identifies shapes by two different IDs: shape IDs and unique IDs. Shape IDs are numeric and uniquely identify shapes within the scope of an individual drawing page or master. They are not unique within the scope of the drawing, however.

Unique IDs are GUIDs. They are unique within the scope of the document.

To convert between shape IDs and unique IDs, you can use two methods of the  **[Page](Visio.Page.md)** object, **[ShapeIDsToUniqueIDs](Visio.Page.ShapeIDsToUniqueIDs.md)** and **[UniqueIDsToShapeIDs](Visio.Page.UniqueIDsToShapeIDs.md)**.

By default, a shape does not have a unique ID. A shape acquires a unique ID only if you get its read-only  **UniqueID** property value by calling the property on the shape, passing it the **visGetOrMake** constant from the **[VisUniqueIDArgs](Visio.visuniqueidargs.md)** enumeration.

If a  **Shape** object has a unique ID, no other shape in the same document will have the same ID.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]