---
title: ODSOFilters.Item Method (Office)
keywords: vbaof11.chm241003
f1_keywords:
- vbaof11.chm241003
ms.prod: office
api_name:
- Office.ODSOFilters.Item
ms.assetid: eff21bc3-dc55-82a4-d405-2d4842c8bfa0
ms.date: 06/08/2017
---


# ODSOFilters.Item Method (Office)

Represents a  **ODSOFilter** object in the **ODSOFilters** collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ A variable that represents an [ODSOFilters](./Office.ODSOFilters.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The number of the item.|

## Return value

Object


## Example

The following example retrieves an  **ODSOFilter** object from the **ODSOFilters** collection.


```vb
oOdsoFilter = oOdsoFilters.Item(1)
```


## See also


[ODSOFilters Object](Office.ODSOFilters.md)



[ODSOFilters Object Members](./overview/Library-Reference/odsofilters-members-office.md)

