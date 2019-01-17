---
title: Application.DeleteCustomList method (Excel)
keywords: vbaxl10.chm133117
f1_keywords:
- vbaxl10.chm133117
ms.prod: excel
api_name:
- Excel.Application.DeleteCustomList
ms.assetid: 41a936f7-05b5-520f-f5c5-172a5ea124d9
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DeleteCustomList method (Excel)

Deletes a custom list.


## Syntax

_expression_. `DeleteCustomList`( `_ListNum_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ListNum_|Required| **Long**|The custom list number. This number must be greater than or equal to 5 (Microsoft Excel has four built-in custom lists that cannot be deleted).|

## Remarks

This method generates an error if the list number is less than 5 or if there's no matching custom list.


## Example

This example deletes a custom list.


```vb
n = Application.GetCustomListNum(Array("cogs", "sprockets", _ 
 "widgets", "gizmos")) 
Application.DeleteCustomList n
```


## See also


[Application Object](Excel.Application(object).md)

