---
title: PivotField.MemoryUsed property (Excel)
keywords: vbaxl10.chm240109
f1_keywords:
- vbaxl10.chm240109
ms.prod: excel
api_name:
- Excel.PivotField.MemoryUsed
ms.assetid: 8faeb893-e0a0-39ed-aa78-4b2b5bb67d69
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.MemoryUsed property (Excel)

Returns the amount of memory currently being used by the object, in bytes. Read-only **Long**.


## Syntax

_expression_.**MemoryUsed**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example displays a message box showing the number of bytes that Microsoft Excel is currently using.

```vb
MsgBox "Microsoft Excel is currently using " & _ 
 Application.MemoryUsed & " bytes"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]