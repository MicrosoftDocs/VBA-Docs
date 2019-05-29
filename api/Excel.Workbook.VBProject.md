---
title: Workbook.VBProject property (Excel)
keywords: vbaxl10.chm199181
f1_keywords:
- vbaxl10.chm199181
ms.prod: excel
api_name:
- Excel.Workbook.VBProject
ms.assetid: 1bef5b7e-e169-fa4b-9810-6cd87ecd0a8d
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.VBProject property (Excel)

Returns a **VBProject** object that represents the Visual Basic project in the specified workbook. Read-only.


## Syntax

_expression_.**VBProject**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example changes the name of the Visual Basic project in the workbook.

```vb
ThisWorkbook.VBProject.Name = "TestProject"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
