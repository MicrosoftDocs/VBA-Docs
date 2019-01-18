---
title: CustomViews.Add method (Excel)
keywords: vbaxl10.chm506075
f1_keywords:
- vbaxl10.chm506075
ms.prod: excel
api_name:
- Excel.CustomViews.Add
ms.assetid: 134d9969-048b-6a53-4f2c-cc83589c5a70
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomViews.Add method (Excel)

Creates a new custom view.


## Syntax

_expression_. `Add`( `_ViewName_` , `_PrintSettings_` , `_RowColSettings_` )

_expression_ A variable that represents a [CustomViews](Excel.CustomViews.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ViewName_|Required| **String**|The name of the new view.|
| _PrintSettings_|Optional| **Variant**| **True** to include print settings in the custom view.|
| _RowColSettings_|Optional| **Variant**| **True** to include settings for hidden rows and columns (including filter information) in the custom view.|

## Return value

A  **[CustomView](Excel.CustomView.md)** object that represents the new custom view.


## Example

This example creates a new custom view named "Summary" in the active workbook.


```vb
ActiveWorkbook.CustomViews.Add "Summary", True, True
```


## See also


[CustomViews Object](Excel.CustomViews.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]