---
title: Worksheet.Select method (Excel)
keywords: vbaxl10.chm174095
f1_keywords:
- vbaxl10.chm174095
ms.prod: excel
api_name:
- Excel.Worksheet.Select
ms.assetid: 2010145e-d36f-7d2b-cfbf-8419c15b31a5
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Select method (Excel)

Selects the object.


## Syntax

_expression_.**Select** (_Replace_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Replace_|Optional| **Variant**| (used only with sheets). **True** to replace the current selection with the specified object. **False** to extend the current selection to include any previously selected objects and the specified object.|

## Remarks

To select a sheet or multiple sheets, use the **Select** method. To make a single sheet the active sheet, use the **[Activate](Excel.Worksheet.Activate(method).md)** method.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
