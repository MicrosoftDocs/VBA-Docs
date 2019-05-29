---
title: Workbook.Post method (Excel)
keywords: vbaxl10.chm199125
f1_keywords:
- vbaxl10.chm199125
ms.prod: excel
api_name:
- Excel.Workbook.Post
ms.assetid: 62ecf3bc-c551-8f06-64cc-a6c141bdf172
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Post method (Excel)

Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.


## Syntax

_expression_.**Post** (_DestName_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestName_|Optional| **Variant**|This argument is ignored. The **Post** method prompts the user to specify the destination for the workbook.|

## Example

This example posts the active workbook.

```vb
ActiveWorkbook.Post
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]