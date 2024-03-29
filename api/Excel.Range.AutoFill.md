---
title: Range.AutoFill method (Excel)
keywords: vbaxl10.chm144083
f1_keywords:
- vbaxl10.chm144083
api_name:
- Excel.Range.AutoFill
ms.assetid: 257f6608-9211-86f9-79de-e3c44df8f3fd
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.AutoFill method (Excel)

Performs an autofill on the cells in the specified range.


## Syntax

_expression_.**AutoFill** (_Destination_, _Type_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Destination_|Required| **Range**|The cells to be filled. The destination must include the source range.|
| _Type_|Optional| **[XlAutoFillType](Excel.XlAutoFillType.md)**|Specifies the fill type.|

## Return value

Variant


## Example

This example performs an autofill on cells A1:A20 on Sheet1, based on the source range A1:A2 on Sheet1. Before running this example, type **1** in cell A1 and type **2** in cell A2.

```vb
Set sourceRange = Worksheets("Sheet1").Range("A1:A2") 
Set fillRange = Worksheets("Sheet1").Range("A1:A20") 
sourceRange.AutoFill Destination:=fillRange
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
