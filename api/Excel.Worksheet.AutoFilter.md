---
title: Worksheet.AutoFilter property (Excel)
keywords: vbaxl10.chm175144
f1_keywords:
- vbaxl10.chm175144
ms.prod: excel
api_name:
- Excel.Worksheet.AutoFilter
ms.assetid: 766f8501-dae7-32a7-9fae-70a87d0a8eba
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.AutoFilter property (Excel)

Returns an **[AutoFilter](excel.autofilter.md)** object if filtering is on. Read-only.


## Syntax

_expression_.**AutoFilter**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

The property returns **Nothing** if filtering is off.

To create an **AutoFilter** object for a worksheet, you must turn autofiltering on for a range on the worksheet either manually or by using the **[AutoFilter](excel.range.autofilter.md)** method of the **Range** object.


## Example

The following example returns AutoFilter for the current worksheet.

```vb
Dim Worksheet1 As Worksheet 
 
Dim returnValue As AutoFilter 
Set returnValue = Worksheet1.AutoFilter
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
