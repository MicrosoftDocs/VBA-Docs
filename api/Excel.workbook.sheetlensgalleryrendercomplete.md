---
title: Workbook.SheetLensGalleryRenderComplete event (Excel)
keywords: vbaxl10.chm503109
f1_keywords:
- vbaxl10.chm503109
ms.prod: excel
ms.assetid: 8ac48e9f-7a15-c674-6d96-e9c1466473bc
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetLensGalleryRenderComplete event (Excel)

Occurs when a callout gallery's icons (dynamic and static) have completed rendering for a worksheet.


## Syntax

_expression_.**SheetLensGalleryRenderComplete** (_Sh_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|A worksheet object.|


## Remarks

This event can be used to indicate when to re-enable screen updating or display additional dialogs, for example.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]