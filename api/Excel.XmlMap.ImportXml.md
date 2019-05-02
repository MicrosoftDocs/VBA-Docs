---
title: XmlMap.ImportXml method (Excel)
keywords: vbaxl10.chm754088
f1_keywords:
- vbaxl10.chm754088
ms.prod: excel
api_name:
- Excel.XmlMap.ImportXml
ms.assetid: 07db07d3-cd0f-08fe-3463-04ca72d084d1
ms.date: 06/08/2017
localization_priority: Normal
---


# XmlMap.ImportXml method (Excel)

Imports XML data from a  **String** variable into cells that have been mapped to the specified **[XmlMap](Excel.XmlMap.md)** object.


## Syntax

_expression_. `ImportXml`( `_XmlData_` , `_Overwrite_` )

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XmlData_|Required| **String**|The string that contains the XML data to import.|
| _Overwrite_|Optional| **Variant**|Specifies whether to overwrite the contents of cells that are currently mapped to the specified XML map. Set to  **True** to overwrite the cells; set to **False** to append the data to the existing range. If this parameter is not specified, the current value of the **[AppendOnImport](Excel.XmlMap.AppendOnImport.md)** property of the XML map determines whether the contents of cells are overwritten or not.|

## Return value

[XlXmlImportResult](Excel.XlXmlImportResult.md)


## Remarks



| **xlXmlImportResult** can be one of the following **xlXmlImportResult** constants.|
| **xlXmlImportElementsTruncated**. The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess**. The XML data file was successfully imported.|
| **xlXmlImportValidationFailed**. The data being imported failed schema validation, but was imported anyway.|

To import the contents of an XML data file into cells mapped to a specific schema map, use the  **[Import](Excel.XmlMap.Import.md)** method of the **XmlMap** object.

If either of the following conditions is true, a run-time error will occur. If more than one condition is true, Excel returns a run-time error for the most severe (they are listed below with the most severe listed first):


- If the XML data contains syntactical errors.
    
- If import is cancelled because not all of the data could fit in the worksheet.
    

## See also


[XmlMap Object](Excel.XmlMap.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]