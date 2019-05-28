---
title: Workbook.IconSets property (Excel)
keywords: vbaxl10.chm199261
f1_keywords:
- vbaxl10.chm199261
ms.prod: excel
api_name:
- Excel.Workbook.IconSets
ms.assetid: c837d2a8-d21d-7432-a409-f49426368556
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.IconSets property (Excel)

This property is used to filter data in a workbook based on a cell icon from the **[IconSets](excel.iconsets.md)** collection. Read-only.


## Syntax

_expression_.**IconSets**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

In the following example, data is filtered by a cell icon.

```vb
Selection.AutoFilter Field:=1, Criteria1:=ActiveWorkbook.IconSets(xl3Arrows).Item(1), Operator:=xlFilterIcon
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]