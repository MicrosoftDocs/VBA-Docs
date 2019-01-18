---
title: Application.NetworkTemplatesPath property (Excel)
keywords: vbaxl10.chm133173
f1_keywords:
- vbaxl10.chm133173
ms.prod: excel
api_name:
- Excel.Application.NetworkTemplatesPath
ms.assetid: 4710091a-a655-dd49-7ad8-0f4c64eda13a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NetworkTemplatesPath property (Excel)

Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. Read-only  **String**.


## Syntax

_expression_. `NetworkTemplatesPath`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example displays the network path where templates are stored.


```vb
Msgbox Application.NetworkTemplatesPath
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]