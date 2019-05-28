---
title: Workbook.XmlImportXml method (Excel)
keywords: vbaxl10.chm199231
f1_keywords:
- vbaxl10.chm199231
ms.prod: excel
api_name:
- Excel.Workbook.XmlImportXml
ms.assetid: b0edbe49-f578-ead0-8371-0196f5c515d4
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.XmlImportXml method (Excel)

Imports an XML data stream that has been previously loaded into memory. Excel uses the first qualifying map found, or if the destination range is specified, Excel automatically lists the data.


## Syntax

_expression_.**XmlImportXml** (_Data_, _ImportMap_, _Overwrite_, _Destination_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Data_|Required| **String**|The data to import.|
| _ImportMap_|Required| **[XmlMap](Excel.XmlMap.md)**|The schema map to apply when importing the file.|
| _Overwrite_|Optional| **Variant**|If a value is not specified for the _Destination_ parameter, this parameter specifies whether to overwrite data that has been mapped to the schema map specified in the _ImportMap_ parameter. Set to **True** to overwrite the data or **False** to append the new data to the existing data. The default value is **True**. <br/><br/>If a value is specified for the _Destination_ parameter, this parameter specifies whether to overwrite existing data. Set to **True** to overwrite existing data or **False** to cancel the import if data would be overwritten. The default value is **True**.|
| _Destination_|Optional| **Variant**|Specifies the range where the list will be created. Excel only uses the top-left corner of the range.|

## Return value

**[XlXmlImportResult](Excel.XlXmlImportResult.md)**


## Remarks

Don't specify a value for the _Destination_ parameter if you want to import data into an existing mapping.

The following conditions cause the **[XmlImport](Excel.Workbook.XmlImport.md)** method to generate run-time errors:

- The specified XML data contains syntax errors.
    
- The import process was cancelled because the specified data cannot fit into the worksheet.
    
- If no qualifying maps are found and the destination range was not specified.
    
Use the **XMLImport** method of the **Workbook** object to import an XML data file into the current workbook.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]