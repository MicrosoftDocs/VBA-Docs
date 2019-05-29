---
title: Workbook.UnprotectSharing method (Excel)
keywords: vbaxl10.chm199158
f1_keywords:
- vbaxl10.chm199158
ms.prod: excel
api_name:
- Excel.Workbook.UnprotectSharing
ms.assetid: edce1744-0906-4b4e-8b98-5d1125047bff
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.UnprotectSharing method (Excel)

Turns off protection for sharing and saves the workbook.


## Syntax

_expression_.**UnprotectSharing** (_SharingPassword_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SharingPassword_|Optional| **Variant**|The workbook password.|

## Example

This example turns off protection for sharing and saves the active workbook.

```vb
ActiveWorkbook.UnprotectSharing
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]