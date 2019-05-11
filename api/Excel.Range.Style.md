---
title: Range.Style property (Excel)
keywords: vbaxl10.chm144204
f1_keywords:
- vbaxl10.chm144204
ms.prod: excel
api_name:
- Excel.Range.Style
ms.assetid: 78c536c9-7fda-3171-2a93-5c4e57bb8207
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Style property (Excel)

Returns or sets a **Variant** value containing a **[Style](Excel.Style.md)** object that represents the style of the specified range.


## Syntax

_expression_.**Style**

_expression_ A variable that represents a **[Range](Excel.Range(object).md)** object.


## Example

This example applies the Normal style to cell A1 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1").Style = "Normal"

```

<br/>

An alternative is the following.

```vb
Worksheets("Sheet1").Range("A1").Style = ThisWorkbook.Styles("Normal")
```

<br/>

If cell B4 on Sheet1 currently has the Normal style applied, this example applies the Percent style.

```vb
If Worksheets("Sheet1").Range("B4").Style = "Normal" Then 
 Worksheets("Sheet1").Range("B4").Style = "Percent" 
End If

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
