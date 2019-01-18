---
title: Workbooks.OpenXML method (Excel)
keywords: vbaxl10.chm203088
f1_keywords:
- vbaxl10.chm203088
ms.prod: excel
api_name:
- Excel.Workbooks.OpenXML
ms.assetid: c16a7842-19e9-6731-146e-038322c248ba
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbooks.OpenXML method (Excel)

Opens an XML data file. Returns a  **[Workbook](Excel.Workbook.md)** object.


## Syntax

_expression_. `OpenXML`( `_Filename_` , `_Stylesheets_` , `_LoadOption_` )

_expression_ A variable that represents a [Workbooks](./Excel.Workbooks.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the file to open.|
| _Stylesheets_|Optional| **Variant**|Either a single value or an array of values that specify which XSL Transformation (XSLT) stylesheet processing instructions to apply.|
| _LoadOption_|Optional| **Variant**|Specifies how Excel opens the XML data file. Can be one of the  **[xlXmlLoadOption](Excel.XlXmlLoadOption.md)** constants.|

## Return value

Workbook


## Remarks





| **xlXmlLoadOption** can be one of these **xlXmlLoadOption** constants.|
| **xlXmlLoadImportToList** Automatically creates an XML List and imports data into the list.|
| **xlXmlLoadMapXml** Loads the XML file into the **XML Source** task pane.|
| **xlXmlLoadOpenXml** Open XML files in the same way that Excel 2002 opens XML files (for backwards compatibility only).|
| **xlXmlLoadPromptUser** Prompts the user and lets them choose the Import method.|

## Example

The following code opens the XML data file "customers.xml" and displays the file's contents in an XML list.


```vb
Sub UseOpenXML() 
 Application.Workbooks.OpenXML _ 
 Filename:="customers.xml", _ 
 LoadOption:=xlXmlLoadImportToList 
End Sub
```


## See also


[Workbooks Object](Excel.Workbooks.md)

