---
title: Graphic.Filename property (Excel)
keywords: vbaxl10.chm694080
f1_keywords:
- vbaxl10.chm694080
api_name:
- Excel.Graphic.Filename
ms.assetid: 8657c279-2c17-57ea-e898-aab0b7b705b4
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Graphic.Filename property (Excel)

Returns or sets the URL (on the intranet or the web) or path (local or network) to the location where the specified source object was saved. Read/write **String**.


## Syntax

_expression_.**Filename**

_expression_ A variable that represents a **[Graphic](Excel.Graphic.md)** object.


## Remarks

The **FileName** property generates an error if a folder in the specified path doesn't exist.


## Example

This example sets the location where the first item in the active workbook is to be saved.

```vb
ActiveWorkbook.PublishObjects(1).FileName = _ 
 "\\Server2\Q1\StockReport.htm"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]