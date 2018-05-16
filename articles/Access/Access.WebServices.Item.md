---
title: WebServices.Item Property (Access)
keywords: vbaac10.chm14553
f1_keywords:
- vbaac10.chm14553
ms.prod: access
api_name:
- Access.WebServices.Item
ms.assetid: 410eb3be-2336-907a-7284-1311e09bb77b
ms.date: 06/08/2017
---


# WebServices.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents a **WebServices** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**||

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


[WebServices Collection](Access.WebServices.md)

