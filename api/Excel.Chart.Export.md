---
title: Chart.Export method (Excel)
keywords: vbaxl10.chm149163
f1_keywords:
- vbaxl10.chm149163
ms.prod: excel
api_name:
- Excel.Chart.Export
ms.assetid: 4dc7dea6-9be8-ccd4-8198-7726b8fad024
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Export method (Excel)

Exports the chart in a graphic format.


## Syntax

_expression_.**Export** (_FileName_, _FilterName_, _Interactive_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the exported file.|
| _FilterName_|Optional| **Variant**|The language-independent name of the graphic filter as it appears in the registry.|
| _Interactive_|Optional| **Variant**| **True** to display the dialog box that contains the filter-specific options. If this argument is **False**, Microsoft Excel uses the default values for the filter. The default value is **False**.|

## Return value

Boolean


## Example

This example exports chart one as a GIF file.

```vb
Worksheets("Sheet1").ChartObjects(1) _ 
.Chart. Export _ 
 FileName:="current_sales.gif", FilterName:="GIF"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
