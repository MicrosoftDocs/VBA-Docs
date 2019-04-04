---
title: Application.NetworkTemplatesPath property (Excel)
keywords: vbaxl10.chm133173
f1_keywords:
- vbaxl10.chm133173
ms.prod: excel
api_name:
- Excel.Application.NetworkTemplatesPath
ms.assetid: 4710091a-a655-dd49-7ad8-0f4c64eda13a
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.NetworkTemplatesPath property (Excel)

Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. Read-only **String**.


## Syntax

_expression_.**NetworkTemplatesPath**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the network path where templates are stored.

```vb
Msgbox Application.NetworkTemplatesPath
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]