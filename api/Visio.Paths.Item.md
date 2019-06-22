---
title: Paths.Item property (Visio)
keywords: vis_sdr.chm15313765
f1_keywords:
- vis_sdr.chm15313765
ms.prod: visio
api_name:
- Visio.Paths.Item
ms.assetid: 85132486-5baa-d3ab-995d-62cf51d4b1da
ms.date: 06/08/2017
localization_priority: Normal
---


# Paths.Item property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections, and for the **Path** and **Selection** objects. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Paths](Visio.Paths.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

Path


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index )
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]