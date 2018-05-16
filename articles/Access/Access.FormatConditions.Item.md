---
title: FormatConditions.Item Property (Access)
keywords: vbaac10.chm10069
f1_keywords:
- vbaac10.chm10069
ms.prod: access
api_name:
- Access.FormatConditions.Item
ms.assetid: 8de3c1bf-86d6-be3f-e4b0-bc4140eee1d3
ms.date: 06/08/2017
---


# FormatConditions.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **FormatCondition**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents a **FormatConditions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the  _expression_ argument. If a numeric expression, the _index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, the _index_ argument must be the name of a member of the collection|

## Remarks

If the value provided for the  _index_ argument doesn't match any existing member of the collection, an error occurs.

The  **Item** property is the default member of a collection, so you don't have to specify it explicitly. For example, the following two lines of code are equivalent:




```vb
Debug.Print Modules(0)
```




```vb
Debug.Print Modules.Item(0)
```


## See also


#### Concepts


[FormatConditions Collection](Access.FormatConditions.md)

