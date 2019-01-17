---
title: Workbook.WebOptions property (Excel)
keywords: vbaxl10.chm199188
f1_keywords:
- vbaxl10.chm199188
ms.prod: excel
api_name:
- Excel.Workbook.WebOptions
ms.assetid: 801742a2-f5d8-5311-ea24-fd428532ba80
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.WebOptions property (Excel)

Returns the  **[WebOptions](Excel.WebOptions.md)** collection, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. Read-only.


## Syntax

_expression_. `WebOptions`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Example

This example specifies that cascading style sheets and Western document encoding be used when items in the first workbook are saved to a Web page.


```vb
Set objWO = Workbooks(1).WebOptions 
objWO.RelyOnCSS = True 
objWO.Encoding = msoEncodingWestern
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]