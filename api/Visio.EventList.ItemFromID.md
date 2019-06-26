---
title: EventList.ItemFromID property (Visio)
keywords: vis_sdr.chm12713775
f1_keywords:
- vis_sdr.chm12713775
ms.prod: visio
api_name:
- Visio.EventList.ItemFromID
ms.assetid: 0a5e73f7-e876-90e1-4b33-15a2593bb9cc
ms.date: 06/08/2017
localization_priority: Normal
---


# EventList.ItemFromID property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

_expression_.**ItemFromID** (_nID_)

_expression_ A variable that represents an **[EventList](Visio.EventList.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

## Return value

Event


## Remarks

The ID of a **Shape** object uniquely identifies the shape within its page or master.

The ID of a **Style** object uniquely identifies the style within its document.

The ID of a **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]