---
title: PivotField.AutoShow method (Excel)
keywords: vbaxl10.chm240112
f1_keywords:
- vbaxl10.chm240112
ms.prod: excel
api_name:
- Excel.PivotField.AutoShow
ms.assetid: 8caea6de-8872-c474-38bd-8d6d78d9f0cc
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.AutoShow method (Excel)

Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report.


## Syntax

_expression_.**AutoShow** (_Type_, _Range_, _Count_, _Field_)

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Long**|Use the **xlAutomatic** [constant](excel.constants.md) to cause the specified PivotTable report to show the items that match the specified criteria. Use **xlManual** to disable this feature.|
| _Range_|Required| **Long**|The location at which to start showing items. Can be either of the following constants: **xlTop** or **xlBottom**.|
| _Count_|Required| **Long**|The number of items to be shown.|
| _Field_|Required| **String**|The name of the base data field. You must specify the unique name (as returned from the **[SourceName](Excel.PivotField.SourceName.md)** property), and not the displayed name.|

## Example

This example shows only the top two companies, based on the sum of sales.

```vb
ActiveSheet.PivotTables("Pivot1").PivotFields("Company") _ 
 .AutoShow xlAutomatic, xlTop, 2, "Sum of Sales"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]