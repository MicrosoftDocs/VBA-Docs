---
title: FillFormat.TwoColorGradient Method (Excel)
keywords: vbaxl10.chm115008
f1_keywords:
- vbaxl10.chm115008
ms.prod: excel
api_name:
- Excel.FillFormat.TwoColorGradient
ms.assetid: 52b66d42-3489-365a-7c9e-368c27f45488
ms.date: 06/08/2017
---


# FillFormat.TwoColorGradient Method (Excel)

Sets the specified fill to a two-color gradient.


## Syntax

 _expression_. `TwoColorGradient`( `_Style_` , `_Variant_` )

 _expression_ A variable that represents a [FillFormat](./Excel.FillFormat.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[MsoGradientStyle](./Office.MsoGradientStyle.md)**|The gradient style.|
| _Variant_|Required| **Integer**|The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If _Style_ is **msoGradientFromCenter** , the _Variant_ argument can only be 1 or 2.|

## See also


[FillFormat Object](Excel.FillFormat.md)

