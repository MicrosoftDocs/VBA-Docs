---
title: Workbook.FollowHyperlink method (Excel)
keywords: vbaxl10.chm199182
f1_keywords:
- vbaxl10.chm199182
ms.prod: excel
api_name:
- Excel.Workbook.FollowHyperlink
ms.assetid: d070ecc9-fbb6-c146-f250-5c99b09063ec
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.FollowHyperlink method (Excel)

Displays a cached document if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.


## Syntax

_expression_.**FollowHyperlink** (_Address_, _SubAddress_, _NewWindow_, _AddHistory_, _ExtraInfo_, _Method_, _HeaderInfo_)

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Address_|Required| **String**|The address of the target document.|
| _SubAddress_|Optional| **Variant**|The location within the target document. The default value is the empty string.|
| _NewWindow_|Optional| **Variant**| **True** to display the target application in a new window. The default value is **False**.|
| _AddHistory_|Optional| **Variant**|Not used. Reserved for future use.|
| _ExtraInfo_|Optional| **Variant**|A **String** or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use _ExtraInfo_ to specify the coordinates of an image map, the contents of a form, or a FAT file name.|
| _Method_|Optional| **Variant**| Specifies the way _ExtraInfo_ is attached. Can be one of the **[MsoExtraInfoMethod](Office.MsoExtraInfoMethod.md)** constants: **msoMethodGet** or **msoMethodPost**.|
| _HeaderInfo_|Optional| **Variant**|A **String** that specifies header information for the HTTP request. The default value is an empty string.|


## Example

This example loads the document at example.microsoft.com in a new browser window and adds it to the History folder.

```vb
ActiveWorkbook.FollowHyperlink Address:="https://example.microsoft.com"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
