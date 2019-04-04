---
title: Application.AutoPercentEntry property (Excel)
keywords: vbaxl10.chm133250
f1_keywords:
- vbaxl10.chm133250
ms.prod: excel
api_name:
- Excel.Application.AutoPercentEntry
ms.assetid: 80ade0a1-84ae-5a17-6a75-189c0c06843d
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AutoPercentEntry property (Excel)

**True** if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered. Read/write **Boolean**.


## Syntax

_expression_.**AutoPercentEntry**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example enables automatic multiplication by 100 for subsequent entries in cells formatted as percentages.

```vb
Application.AutoPercentEntry = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]