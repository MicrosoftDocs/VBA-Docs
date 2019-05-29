---
title: Workbook.ExportAsFixedFormat method (Excel)
keywords: vbaxl10.chm199260
f1_keywords:
- vbaxl10.chm199260
ms.prod: excel
api_name:
- Excel.Workbook.ExportAsFixedFormat
ms.assetid: 4d72247c-bab9-3475-4792-8899c959393c
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ExportAsFixedFormat method (Excel)

The **ExportAsFixedFormat** method is used to publish a workbook to either the PDF or XPS format.


## Syntax

_expression_.**ExportAsFixedFormat** (_Type_, _FileName_, _Quality_, _IncludeDocProperties_, _IgnorePrintAreas_, _From_, _To_, _OpenAfterPublish_, _FixedFormatExtClassPtr_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFixedFormatType](excel.xlfixedformattype.md)**|Can be either **xlTypePDF** or **xlTypeXPS**.|
| _FileName_|Optional| **Variant**|A string that indicates the name of the file to be saved. You can include a full path, or Excel saves the file in the current folder.|
| _Quality_|Optional| **Variant**|Can be set to either of the following **[XlFixedFormatQuality](Excel.XlFixedFormatQuality.md)** constants: **xlQualityStandard** or **xlQualityMinimum**.|
| _IncludeDocProperties_|Optional| **Variant**|Set to **True** to indicate that document properties should be included, or set to **False** to indicate that they are omitted.|
| _IgnorePrintAreas_|Optional| **Variant**|If set to **True**, ignores any print areas set when publishing. If set to **False**, uses the print areas set when publishing.|
| _From_|Optional| **Variant**|The number of the page at which to start publishing. If this argument is omitted, publishing starts at the beginning.|
| _To_|Optional| **Variant**|The number of the last page to publish. If this argument is omitted, publishing ends with the last page.|
| _OpenAfterPublish_|Optional| **Variant**|If set to **True**, displays the file in the viewer after it is published. If set to **False**, the file is published but not displayed.|
| _FixedFormatExtClassPtr_|Optional| **Variant**|Pointer to the **FixedFormatExt** class.|

## Example

The following example creates the PDF at standard quality in the current file's directory and displays the file in the viewer after it is published. An error occurs if the PDF add-in is not currently installed.

```vb
ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF FileName:="sales.pdf" Quality:=xlQualityStandard OpenAfterPublish:=True 
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
