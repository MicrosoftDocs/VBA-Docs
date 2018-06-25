---
title: ODSOColumns.Item Method (Office)
keywords: vbaof11.chm234003
f1_keywords:
- vbaof11.chm234003
ms.prod: office
api_name:
- Office.ODSOColumns.Item
ms.assetid: be6035d4-aac3-879d-ab87-2aa57a70756c
ms.date: 06/08/2017
---


# ODSOColumns.Item Method (Office)

Specifies an  **ODSOColumn** object in the **ODSOColumns** collection.


## Syntax

 _expression_. `Item`( `_varIndex_` )

 _expression_ A variable that represents an [ODSOColumns](./Office.ODSOColumns.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _varIndex_|Required|**Variant**|The index number of the item.|

### Return Value

Object


## Example

The following example retrieves an  **ODSOColumn** object from the **ODSOColumns** collection.


```vb
oOdsoColumn = oOdsoColumns.Item(2)
```


## See also


[ODSOColumns Object](Office.ODSOColumns.md)
#### Other resources


[ODSOColumns Object Members](./overview/odsocolumns-members-office.md)

