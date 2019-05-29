---
title: Worksheet.EnablePivotTable property (Excel)
keywords: vbaxl10.chm175097
f1_keywords:
- vbaxl10.chm175097
ms.prod: excel
api_name:
- Excel.Worksheet.EnablePivotTable
ms.assetid: 8cd09896-9752-677f-a7fd-da46d68ac42a
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.EnablePivotTable property (Excel)

**True** if PivotTable controls and actions are enabled when user-interface-only protection is turned on. Read/write **Boolean**.


## Syntax

_expression_.**EnablePivotTable**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

This property applies to each worksheet and isn't saved with the worksheet or session.

There must be a sufficient number of unlocked cells below and to the right of the PivotTable report for Microsoft Excel to recalculate and display the PivotTable report.


## Example

This example enables PivotTable controls on a protected worksheet.

```vb
ActiveSheet.EnablePivotTable = True 
ActiveSheet.Protect contents:=True, userInterfaceOnly:=True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]