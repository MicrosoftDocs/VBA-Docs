---
title: Application.MergeInstances property (Excel)
keywords: vbaxl10.chm133343
f1_keywords:
- vbaxl10.chm133343
ms.assetid: f406f2b2-802e-421c-9a80-f6f96a7b7c28
ms.date: 04/05/2019
ms.prod: excel
localization_priority: Normal
---


# Application.MergeInstances property (Excel)

**True** to merge multiple instances of the application into a single instance. Read/write **Boolean**.


## Syntax

_expression_.**MergeInstances**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example causes multiple instances of the application to be merged into a single instance.

```vb
Application.MergeInstances = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]