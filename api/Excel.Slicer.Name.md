---
title: Slicer.Name property (Excel)
keywords: vbaxl10.chm905073
f1_keywords:
- vbaxl10.chm905073
ms.prod: excel
api_name:
- Excel.Slicer.Name
ms.assetid: cc8508d3-82fc-365b-c632-2565fd0071c5
ms.date: 05/16/2019
localization_priority: Normal
---


# Slicer.Name property (Excel)

Returns or sets the name of the specified slicer. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Slicer](Excel.Slicer.md)** object.


## Return value

**String**


## Remarks

The name must be unique across all slicers within a workbook. 

The default name uses the text of the field name of the PivotField on which the slicer is based, and if necessary, appends a space and number to make the name unique.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]