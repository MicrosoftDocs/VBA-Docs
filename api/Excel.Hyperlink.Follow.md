---
title: Hyperlink.Follow method (Excel)
keywords: vbaxl10.chm536082
f1_keywords:
- vbaxl10.chm536082
ms.prod: excel
api_name:
- Excel.Hyperlink.Follow
ms.assetid: cdf02d4c-9987-eaed-061b-0f3813d4204b
ms.date: 04/26/2019
localization_priority: Normal
---


# Hyperlink.Follow method (Excel)

Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.


## Syntax

_expression_.**Follow** (_NewWindow_, _AddHistory_, _ExtraInfo_, _Method_, _HeaderInfo_)

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewWindow_|Optional| **Variant**| **True** to display the target application in a new window. The default value is **False**.|
| _AddHistory_|Optional| **Variant**|Not used. Reserved for future use.|
| _ExtraInfo_|Optional| **Variant**|A **String** or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use _ExtraInfo_ to specify the coordinates of an image map, the contents of a form, or a FAT file name.|
| _Method_|Optional| **Variant**|Specifies the way _ExtraInfo_ is attached. Can be one of the **[MsoExtraInfoMethod](Office.MsoExtraInfoMethod.md)** constants.|
| _HeaderInfo_|Optional| **Variant**|A **String** that specifies header information for the HTTP request. The default value is an empty string.|

## Example

This example loads the document attached to the hyperlink on shape one on worksheet one.

```vb
Worksheets(1).Shapes(1).Hyperlink.Follow NewWindow:=True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
