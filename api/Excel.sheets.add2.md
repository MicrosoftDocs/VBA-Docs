---
title: Sheets.Add2 method (Excel)
keywords: vbaxl10.chm152090
f1_keywords:
- vbaxl10.chm152090
ms.prod: excel
ms.assetid: f44b3ef1-8452-4e26-b91c-d24124fa5bc6
ms.date: 06/08/2017
localization_priority: Normal
---


# Sheets.Add2 method (Excel)

This method is only implemented for the  **Charts** collection object and will produce a run time error if used on the **Sheets** and **Worksheets** objects.


## Syntax

_expression_. `Add2`_(Before,_ _After,_ _Count,_ _NewLayout)_

_expression_ A variable that represents a [Sheets](Excel.Sheets.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|**Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional|**Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional|**Variant**|The number of sheets to be added. The default value is one.|
| _NewLayout_|Optional|**Variant**||

## Return value

 **OBJECT**


## See also


[Sheets Object](Excel.Sheets.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
