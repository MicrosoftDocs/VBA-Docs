---
title: AccelItems.Item property (Visio)
keywords: vis_sdr.chm14613765
f1_keywords:
- vis_sdr.chm14613765
ms.prod: visio
api_name:
- Visio.AccelItems.Item
ms.assetid: c6ac3d03-4b13-141f-d1fd-dfbf671435fd
ms.date: 06/24/2019
localization_priority: Normal
---


# AccelItems.Item property (Visio)

Returns an object from a collection. The **Item** property is the default property for all collections. Read-only.


## Syntax

_expression_.**Item** (_lIndex_)

_expression_ A variable that represents an **[AccelItems](Visio.AccelItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

**[AccelItem](Visio.AccelItem.md)**


## Remarks

When retrieving objects from a collection, you can omit **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax given previously.

```vb
objRet = object(index )
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]