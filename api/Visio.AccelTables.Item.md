---
title: AccelTables.Item Property (Visio)
keywords: vis_sdr.chm14813765
f1_keywords:
- vis_sdr.chm14813765
ms.prod: visio
api_name:
- Visio.AccelTables.Item
ms.assetid: 0dafb64d-fc3b-4b00-2e2d-062fb55216ef
ms.date: 06/08/2017
localization_priority: Normal
---


# AccelTables.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_. `Item`( `_lIndex_` )

 _expression_ A variable that represents a [AccelTables](./Visio.AccelTables.md) collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

AccelTable


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index )
```


