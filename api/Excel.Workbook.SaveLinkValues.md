---
title: Workbook.SaveLinkValues property (Excel)
keywords: vbaxl10.chm199148
f1_keywords:
- vbaxl10.chm199148
ms.prod: excel
api_name:
- Excel.Workbook.SaveLinkValues
ms.assetid: ee69911f-5a4a-5c2b-c14a-cd562f3ba9f4
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SaveLinkValues property (Excel)

**True** if Microsoft Excel saves external link values with the workbook. Read/write **Boolean**.


## Syntax

_expression_.**SaveLinkValues**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example causes Microsoft Excel to save external link values with the active workbook.

```vb
ActiveWorkbook.SaveLinkValues = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]