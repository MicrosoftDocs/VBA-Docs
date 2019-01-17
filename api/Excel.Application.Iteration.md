---
title: Application.Iteration property (Excel)
keywords: vbaxl10.chm133152
f1_keywords:
- vbaxl10.chm133152
ms.prod: excel
api_name:
- Excel.Application.Iteration
ms.assetid: 51e5bd34-844b-3367-951a-6f2f8f9acf90
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Iteration property (Excel)

 **True** if Microsoft Excel will use iteration to resolve circular references. Read/write **Boolean**.


## Syntax

_expression_. `Iteration`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example sets the  **Iteration** property to **True** so that circular references will be resolved by iteration.


```vb
Application.Iteration = True
```


## See also


[Application Object](Excel.Application(object).md)

