---
title: ControlFormat.LockedText property (Excel)
keywords: vbaxl10.chm630084
f1_keywords:
- vbaxl10.chm630084
ms.prod: excel
api_name:
- Excel.ControlFormat.LockedText
ms.assetid: 3b663597-4dec-8e9c-9d85-d07e162c4243
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.LockedText property (Excel)

**True** if the text in the specified object will be locked to prevent changes when the workbook is protected. Read/write **Boolean**.


## Syntax

_expression_.**LockedText**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Example

This example locks text in embedded chart one when the workbook is protected.

```vb
Worksheets(1).ChartObjects(1).LockedText = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]