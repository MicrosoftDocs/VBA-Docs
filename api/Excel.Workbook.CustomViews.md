---
title: Workbook.CustomViews property (Excel)
keywords: vbaxl10.chm199164
f1_keywords:
- vbaxl10.chm199164
ms.prod: excel
api_name:
- Excel.Workbook.CustomViews
ms.assetid: 286f6d5a-fb91-a339-8e74-9014ab7f4835
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.CustomViews property (Excel)

Returns a **[CustomViews](Excel.CustomViews.md)** collection that represents all the custom views for the workbook.


## Syntax

_expression_.**CustomViews**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example creates a new custom view named Summary in the active workbook.

```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]