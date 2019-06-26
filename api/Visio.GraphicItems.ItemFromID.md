---
title: GraphicItems.ItemFromID property (Visio)
keywords: vis_sdr.chm16813775
f1_keywords:
- vis_sdr.chm16813775
ms.prod: visio
api_name:
- Visio.GraphicItems.ItemFromID
ms.assetid: 2d74816f-b667-25f7-7647-ae14e4b8fcad
ms.date: 06/08/2017
localization_priority: Normal
---


# GraphicItems.ItemFromID property (Visio)

Returns a **GraphicItem** object from the **GraphicItems** collection by using the unique ID of the object. Read-only.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**ItemFromID** (_ObjectID_)

_expression_ A variable that represents a **[GraphicItems](Visio.GraphicItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectID_|Required| **Long**|The unique ID of the **GraphicItem** object to retrieve.|

## Return value

GraphicItem


## Remarks

You can get the ID of a **GraphicItem** object by getting the value of the **GraphicItem.ID** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]