---
title: Application.CustomListCount property (Excel)
keywords: vbaxl10.chm133100
f1_keywords:
- vbaxl10.chm133100
ms.prod: excel
api_name:
- Excel.Application.CustomListCount
ms.assetid: 98a32161-e413-a0b7-a6be-4b11ae90fc00
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CustomListCount property (Excel)

Returns the number of defined custom lists (including built-in lists). Read-only **Long**.


## Syntax

_expression_.**CustomListCount**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the number of custom lists that are currently defined.

```vb
MsgBox "There are currently " & Application.CustomListCount & _ 
 " defined custom lists."
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]