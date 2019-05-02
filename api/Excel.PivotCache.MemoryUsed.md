---
title: PivotCache.MemoryUsed property (Excel)
keywords: vbaxl10.chm227077
f1_keywords:
- vbaxl10.chm227077
ms.prod: excel
api_name:
- Excel.PivotCache.MemoryUsed
ms.assetid: f68731ec-053e-79e9-861f-2c225b827e96
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.MemoryUsed property (Excel)

Returns the amount of memory currently being used by the object, in bytes. Read-only **Long**.


## Syntax

_expression_.**MemoryUsed**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

For **PivotCache** objects, this property reflects the transient state of the cache at the time that it's queried.

If the **PivotCache** object has no PivotTable report attached to it, this property returns 0 (zero).


## Example

This example displays a message box showing the number of bytes that Microsoft Excel is currently using.

```vb
MsgBox "Microsoft Excel is currently using " & _ 
 Application.MemoryUsed & " bytes"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]