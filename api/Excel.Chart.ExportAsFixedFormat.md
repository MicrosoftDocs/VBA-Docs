---
title: Chart.ExportAsFixedFormat method (Excel)
keywords: vbaxl10.chm149183
f1_keywords:
- vbaxl10.chm149183
ms.prod: excel
api_name:
- Excel.Chart.ExportAsFixedFormat
ms.assetid: 4fa48315-c9e9-944c-71c3-72ec9894daac
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.ExportAsFixedFormat method (Excel)

Exports to a file of the specified format.


## Syntax

_expression_.**ExportAsFixedFormat** (_Type_, _FileName_, _Quality_, _IncludeDocProperties_, _IgnorePrintAreas_, _From_, _To_, _OpenAfterPublish_, _FixedFormatExtClassPtr_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFixedFormatType](excel.xlfixedformattype.md)**|The type of file format to export to.|
| _FileName_|Optional| **Variant**|The file name of the file to be saved. You can include a full path, or Microsoft Excel saves the file in the current folder.|
| _Quality_|Optional| **Variant**|Optional **[XlFixedFormatQuality](Excel.XlFixedFormatQuality.md)**. Specifies the quality of the published file.|
| _IncludeDocProperties_|Optional| **Variant**| **True** to include the document properties; otherwise, **False**.|
| _IgnorePrintAreas_|Optional| **Variant**| **True** to ignore any print areas set when publishing; otherwise, **False**.|
| _From_|Optional| **Variant**|The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.|
| _To_|Optional| **Variant**|The number of the last page to publish. If this argument is omitted, publishing ends with the last page.|
| _OpenAfterPublish_|Optional| **Variant**| **True** to display the file in the viewer after it is published; otherwise, **False**.|
| _FixedFormatExtClassPtr_|Optional| **Variant**|Pointer to the **FixedFormatExt** class.|

## Remarks

This method also supports initializing an add-in to export a file to a fixed-format file. For example, Excel will perform file format conversion if the converters are present. The conversion is usually initiated by the user.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]