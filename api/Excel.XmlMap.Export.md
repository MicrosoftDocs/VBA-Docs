---
title: XmlMap.Export Method (Excel)
keywords: vbaxl10.chm754089
f1_keywords:
- vbaxl10.chm754089
ms.prod: excel
api_name:
- Excel.XmlMap.Export
ms.assetid: 174f902f-7244-866d-b16c-6a6bcf0ae58b
ms.date: 06/08/2017
---


# XmlMap.Export Method (Excel)

Exports the contents of cells mapped to the specified  **[XmlMap](Excel.XmlMap.md)** object to an XML data file.


## Syntax

 _expression_. `Export`( `_Url_` , `_Overwrite_` )

 _expression_ An expression that returns an [XmlMap](./Excel.XmlMap.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Url_|Required| **String**|The path and file name of the XML data file to export to.|
| _Overwrite_|Optional| **Variant**|Set to  **True** to overwrite the file specified in the _URL_ parameter if the file exists. The default value is **False** .|

### Return value

 **An [XlXmlExportResult](Excel.XlXmlExportResult.md)** value that indicates the result of the method.


## Remarks

Use the  **[ExportXml](Excel.XmlMap.ExportXml.md)** method to export the contents of the mapped cells to a **String** variable.


## See also


[XmlMap Object](Excel.XmlMap.md)

