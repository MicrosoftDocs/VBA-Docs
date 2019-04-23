---
title: Chart.ProtectFormatting property (Excel)
keywords: vbaxl10.chm149157
f1_keywords:
- vbaxl10.chm149157
ms.prod: excel
api_name:
- Excel.Chart.ProtectFormatting
ms.assetid: 71630b7f-6c89-869d-cd5b-d0a7bacd904a
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.ProtectFormatting property (Excel)

**True** if chart formatting cannot be modified by the user. Read/write **Boolean**.


## Syntax

_expression_.**ProtectFormatting**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

This property is not persisted when the file is saved. If you set this property to **True** and then reopen the file, it will no longer be set to **True**.


## Example

This example protects the formatting of embedded chart one on worksheet one.

```vb
Worksheets(1).ChartObjects(1).Chart.ProtectFormatting = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]