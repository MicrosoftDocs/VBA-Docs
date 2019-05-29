---
title: Workbook.ReloadAs method (Excel)
keywords: vbaxl10.chm199189
f1_keywords:
- vbaxl10.chm199189
ms.prod: excel
api_name:
- Excel.Workbook.ReloadAs
ms.assetid: ce6a9d1a-7945-3dca-ff2d-a42289c2ccf9
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ReloadAs method (Excel)

Reloads a workbook based on an HTML document, using the specified document encoding.


## Syntax

_expression_.**ReloadAs** (_Encoding_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Encoding_|Required| **[MsoEncoding](Office.MsoEncoding.md)**|The encoding that is to be applied to the workbook.|

## Remarks

Only **MsoEncoding** constants that are applicable to HTML work with the **ReloadAs** method.


## Example

This example reloads the first workbook, using Western document encoding.

```vb
Workbooks(1).ReloadAs Encoding:=msoEncodingWestern
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]