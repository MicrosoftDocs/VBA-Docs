---
title: Application.CustomListCount property (Excel)
keywords: vbaxl10.chm133100
f1_keywords:
- vbaxl10.chm133100
ms.prod: excel
api_name:
- Excel.Application.CustomListCount
ms.assetid: 98a32161-e413-a0b7-a6be-4b11ae90fc00
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomListCount property (Excel)

Returns the number of defined custom lists (including built-in lists). Read-only  **Long**.


## Syntax

_expression_. `CustomListCount`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example displays the number of custom lists that are currently defined.


```vb
MsgBox "There are currently " & Application.CustomListCount & _ 
 " defined custom lists."
```


## See also


[Application Object](Excel.Application(object).md)

