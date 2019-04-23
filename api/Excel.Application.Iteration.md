---
title: Application.Iteration property (Excel)
keywords: vbaxl10.chm133152
f1_keywords:
- vbaxl10.chm133152
ms.prod: excel
api_name:
- Excel.Application.Iteration
ms.assetid: 51e5bd34-844b-3367-951a-6f2f8f9acf90
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Iteration property (Excel)

**True** if Microsoft Excel uses iteration to resolve circular references. Read/write **Boolean**.


## Syntax

_expression_.**Iteration**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the **Iteration** property to **True** so that circular references are resolved by iteration.

```vb
Application.Iteration = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]