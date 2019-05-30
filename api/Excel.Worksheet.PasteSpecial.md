---
title: Worksheet.PasteSpecial method (Excel)
keywords: vbaxl10.chm175155
f1_keywords:
- vbaxl10.chm175155
ms.prod: excel
api_name:
- Excel.Worksheet.PasteSpecial
ms.assetid: 8fa41a45-e3d1-29e0-3968-877bcfdf4b57
ms.date: 05/30/2019
localization_priority: Normal
---

# Worksheet.PasteSpecial method (Excel)

Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.


## Syntax

_expression_.**PasteSpecial** (_Format_, _Link_, _DisplayAsIcon_, _IconFileName_, _IconIndex_, _IconLabel_, _NoHTMLFormatting_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Format_|Optional| **Variant**|A string that specifies the Clipboard format of the data.|
| _Link_|Optional| **Variant**| **True** to establish a link to the source of the pasted data. If the source data isn't suitable for linking or the source application doesn't support linking, this parameter is ignored. The default value is **False**.|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the pasted data as an icon. The default value is **False**.|
| _IconFileName_|Optional| **Variant**|The name of the file that contains the icon to use if _DisplayAsIcon_ is **True**.|
| _IconIndex_|Optional| **Variant**|The index number of the icon within the icon file.|
| _IconLabel_|Optional| **Variant**|The text label of the icon.|
| _NoHTMLFormatting_|Optional| **Variant**| **True** to remove all formatting, hyperlinks, and images from HTML. **False** to paste HTML as is. The default value is **False**.|

## Remarks

> [!NOTE] 
> _NoHTMLFormatting_ only matters when _Format_ = "HTML"; in all other cases, _NoHTMLFormatting_ is ignored.

You must select the destination range before you use this method.

This method may modify the sheet selection, depending on the contents of the Clipboard.

For developers of languages other than English, you can substitute one of the following constants (0-5) to correspond with the string equivalent of the picture file format.

|Format argument|String equivalent|
|:-----|:-----|
|0|"Picture (PNG)"|
|1|"Picture (JPEG)"|
|2|"Picture (GIF)"|
|3|"Picture (Enhanced Metafile)"|
|4|"Bitmap"|
|5|"Microsoft Office Drawing Object"|

## Example

This example pastes a Microsoft Word document object from the Clipboard to cell D1 on Sheet1.

```vb
Worksheets("Sheet1").Range("D1").Select 
ActiveSheet.PasteSpecial format:= _ 
 "Microsoft Word 8.0 Document Object"
```

<br/>

This example pastes a picture object and does not display it as an icon.

```vb
Worksheets("Sheet1").Range("F5").PasteSpecial _ 
 Format:="Picture (Enhanced Metafile)", Link:=False,
 DisplayAsIcon:=False 

```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
