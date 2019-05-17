---
title: VPageBreak.Location property (Excel)
keywords: vbaxl10.chm156078
f1_keywords:
- vbaxl10.chm156078
ms.prod: excel
api_name:
- Excel.VPageBreak.Location
ms.assetid: d039049f-5b08-d867-c874-f25ca0dbe70f
ms.date: 05/18/2019
localization_priority: Normal
---


# VPageBreak.Location property (Excel)

Returns the cell (a **[Range](Excel.Range(object).md)** object) that defines the page-break location. Vertical page breaks are aligned with the left edge of the location cell. Read-only **Range**.


## Syntax

_expression_.**Location** 

_expression_ A variable that represents a **[VPageBreak](Excel.VPageBreak.md)** object.


## Example

This example stores the vertical page-break location in a **Range** object.

```vb
Dim r as Range
Set r = Worksheets(1).VPageBreaks(1).Location
```

> [!NOTE] 
> **Location** is read-only, and can only be used to return the current vertical page-break location. To change the location of a **VPageBreak**, you must use the **[Dragoff](Excel.VPageBreak.DragOff.md)** method. 



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]