---
title: Workbook.DisplayDrawingObjects property (Excel)
keywords: vbaxl10.chm199098
f1_keywords:
- vbaxl10.chm199098
ms.prod: excel
api_name:
- Excel.Workbook.DisplayDrawingObjects
ms.assetid: 78eec8af-416d-88e6-d1f4-0b97a008f752
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.DisplayDrawingObjects property (Excel)

Returns or sets how shapes are displayed. Read/write **Long**.


## Syntax

_expression_.**DisplayDrawingObjects**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Can be one of the following **[XlDisplayDrawingObjects](excel.xldisplaydrawingobjects.md)** constants: **xlDisplayShapes**, **xlPlaceholders**, or **xlHide**.

## Example

This example hides all the shapes in the active workbook.

```vb
ActiveWorkbook.DisplayDrawingObjects = xlHide
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]