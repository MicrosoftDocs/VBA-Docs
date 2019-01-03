---
title: ODSOColumns.Item method (Office)
keywords: vbaof11.chm234003
f1_keywords:
- vbaof11.chm234003
ms.prod: office
api_name:
- Office.ODSOColumns.Item
ms.assetid: be6035d4-aac3-879d-ab87-2aa57a70756c
ms.date: 06/08/2017
---


# ODSOColumns.Item method (Office)

Specifies an  **ODSOColumn** object in the **ODSOColumns** collection.


## Syntax

_expression_. `Item`( `_varIndex_` )

_expression_ A variable that represents an [ODSOColumns](Office.ODSOColumns.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _varIndex_|Required|**Variant**|The index number of the item.|

## Return value

Object


## Example

The following example retrieves an  **ODSOColumn** object from the **ODSOColumns** collection.


```vb
oOdsoColumn = oOdsoColumns.Item(2)
```


## See also


[ODSOColumns Object](Office.ODSOColumns.md)



[ODSOColumns Object Members](./overview/Library-Reference/odsocolumns-members-office.md)

