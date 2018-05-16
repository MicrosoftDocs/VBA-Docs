---
title: AllFunctions.Item Property (Access)
keywords: vbaac10.chm12680
f1_keywords:
- vbaac10.chm12680
ms.prod: access
api_name:
- Access.AllFunctions.Item
ms.assetid: d705b1b9-5509-4a1e-8a0e-55357a544d62
ms.date: 06/08/2017
---


# AllFunctions.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **AccessObject**.


## Syntax

 _expression_. **Item**( ** _var_** )

 _expression_ A variable that represents an **AllFunctions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _var_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the  _expression_ argument. If a numeric expression, the _index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, the _index_ argument must be the name of a member of the collection|

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


[AllFunctions Collection](Access.AllFunctions.md)

