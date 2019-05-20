---
title: XmlMap.Import method (Excel)
keywords: vbaxl10.chm754087
f1_keywords:
- vbaxl10.chm754087
ms.prod: excel
api_name:
- Excel.XmlMap.Import
ms.assetid: 60265bbd-4994-8fba-7072-ec5dada885d3
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlMap.Import method (Excel)

Imports data from the specified XML data file into cells that have been mapped to the specified **XmlMap** object.


## Syntax

_expression_.**Import** (_Url_, _Overwrite_)

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Url_|Required| **String**|The path to the XML data to import. The path can be specified in Universal Naming convention (UNC) or Uniform Resource Locator (URL) format. The file can be an XML data file.|
| _Overwrite_|Optional| **Variant**|Set to **True** to overwrite existing data. Set to **False** to append to existing data. The default value is **False**.|

## Return value

An **[XlXmlImportResult](Excel.XlXmlImportResult.md)** value that indicates the result of the method.


## Remarks

If either of the following conditions is **True**, a run-time error occurs. If more than one condition is **True**, Excel returns a run-time error for the most severe (they are listed with the most severe listed first):

- If the XML data contains syntactical errors. 
- If the import is cancelled because not all the data could fit on the worksheet.
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]