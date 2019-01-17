---
title: Application.MergeInstances property (Excel)
keywords: vbaxl10.chm133343
f1_keywords:
- vbaxl10.chm133343
ms.assetid: f406f2b2-802e-421c-9a80-f6f96a7b7c28
ms.date: 06/08/2017
ms.prod: excel
localization_priority: Normal
---


# Application.MergeInstances property (Excel)

 **True** to merge multiple instances of the application into a single instance. Read/Write **Boolean**.


## Syntax

_expression_. `MergeInstances`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example causes multiple instances of the application to be merged into a single instance.


```vb
Application.MergeInstances = True
```


## See also


[Application Object](Excel.Application(object).md)

