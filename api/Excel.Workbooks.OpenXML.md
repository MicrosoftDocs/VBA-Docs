---
title: Workbooks.OpenXML method (Excel)
keywords: vbaxl10.chm203088
f1_keywords:
- vbaxl10.chm203088
ms.prod: excel
api_name:
- Excel.Workbooks.OpenXML
ms.assetid: c16a7842-19e9-6731-146e-038322c248ba
ms.date: 05/18/2019
localization_priority: Normal
---


# Workbooks.OpenXML method (Excel)

Opens an XML data file. Returns a **[Workbook](Excel.Workbook.md)** object.


## Syntax

_expression_.**OpenXML** (_FileName_, _Stylesheets_, _LoadOption_)

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to open.|
| _Stylesheets_|Optional| **Variant**|Either a single value or an array of values that specify which XSL Transformation (XSLT) stylesheet processing instructions to apply.|
| _LoadOption_|Optional| **Variant**|Specifies how Excel opens the XML data file. Can be one of the **[XlXmlLoadOption](Excel.XlXmlLoadOption.md)** constants.|

## Return value

Workbook

## Example

The following code opens the XML data file Customers.xml and displays the file's contents in an XML list.

```vb
Sub UseOpenXML() 
 Application.Workbooks.OpenXML _ 
 Filename:="Customers.xml", _ 
 LoadOption:=xlXmlLoadImportToList 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
