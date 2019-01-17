---
title: Application.AlertBeforeOverwriting property (Excel)
keywords: vbaxl10.chm133077
f1_keywords:
- vbaxl10.chm133077
ms.prod: excel
api_name:
- Excel.Application.AlertBeforeOverwriting
ms.assetid: 75c69d9d-bd6e-c0c9-71c4-c9d92333d233
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AlertBeforeOverwriting property (Excel)

 **True** if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation. Read/write **Boolean**.


## Syntax

_expression_. `AlertBeforeOverwriting`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example causes Microsoft Excel to display an alert before overwriting nonblank cells during drag-and-drop editing.


```vb
Application.AlertBeforeOverwriting = True
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]