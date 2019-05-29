---
title: Workbook.Styles property (Excel)
keywords: vbaxl10.chm199154
f1_keywords:
- vbaxl10.chm199154
ms.prod: excel
api_name:
- Excel.Workbook.Styles
ms.assetid: c9a70be9-cab5-ea5f-2e3f-949b1acf43d9
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Styles property (Excel)

Returns a **[Styles](Excel.Styles.md)** collection that represents all the styles in the specified workbook. Read-only.


## Syntax

_expression_.**Styles**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example deletes the user-defined style Stock Quote Style from the active workbook.

```vb
ActiveWorkbook.Styles("Stock Quote Style").Delete
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]