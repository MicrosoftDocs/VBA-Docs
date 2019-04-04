---
title: Application.CalculateBeforeSave property (Excel)
keywords: vbaxl10.chm133083
f1_keywords:
- vbaxl10.chm133083
ms.prod: excel
api_name:
- Excel.Application.CalculateBeforeSave
ms.assetid: 133dbe08-8f41-c07c-8362-48412ed7c086
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CalculateBeforeSave property (Excel)

**True** if workbooks are calculated before they're saved to disk (if the **[Calculation](Excel.Application.Calculation.md)** property is set to **xlManual**). This property is preserved even if you change the **Calculation** property. Read/write **Boolean**.


## Syntax

_expression_.**CalculateBeforeSave**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets Microsoft Excel to calculate workbooks before they're saved to disk.

```vb
Application.Calculation = xlManual 
Application.CalculateBeforeSave = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]