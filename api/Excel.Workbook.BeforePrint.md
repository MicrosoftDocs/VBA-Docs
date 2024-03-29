---
title: Workbook.BeforePrint event (Excel)
keywords: vbaxl10.chm503078
f1_keywords:
- vbaxl10.chm503078
api_name:
- Excel.Workbook.BeforePrint
ms.assetid: 2c97cb32-2bb3-2848-b5ed-32d9129af080
ms.date: 05/29/2019
ms.localizationpriority: medium
---


# Workbook.BeforePrint event (Excel)

Occurs before the workbook (or anything in it) is printed.


## Syntax

_expression_.**BeforePrint** (_Cancel_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the workbook isn't printed when the procedure is finished.|

## Return value

**Nothing**


## Example

This example recalculates all worksheets in the active workbook before printing anything.

```vb
Private Sub Workbook_BeforePrint(Cancel As Boolean) 
 For Each wk in Worksheets 
 wk.Calculate 
 Next 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
