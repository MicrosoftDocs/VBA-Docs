---
title: Formatting and VBA codes for headers and footers
keywords: vbaxl10.chm5201409
f1_keywords:
- vbaxl10.chm5201409
ms.prod: excel
ms.assetid: 70013db6-bb60-8c19-5ef4-1cb54f79b68c
ms.date: 06/08/2017
localization_priority: Normal
---

# Formatting and VBA codes for headers and footers

The following special formatting and Visual Basic for Applications (VBA) codes can be included as a part of the header and footer properties (**[LeftHeader](../../../api/Excel.Page.LeftHeader.md)**, **[CenterHeader](../../../api/Excel.Page.CenterHeader.md)**, **[RightHeader](../../../api/Excel.PageSetup.RightHeader.md)**, **[LeftFooter](../../../api/Excel.Page.LeftFooter.md)**, **[CenterFooter](../../../api/Excel.Page.CenterFooter.md)**, and **[RightFooter](../../../api/Excel.Page.RightFooter.md)**).

<br/>

|**Format code**|**Description**|
|:-----|:-----|
|&L|Left aligns the characters that follow.|
|&C|Centers the characters that follow.|
|&R|Right aligns the characters that follow.|
|&E|Turns double-underline printing on or off.|
|&X|Turns superscript printing on or off.|
|&Y|Turns subscript printing on or off.|
|&B|Turns bold printing on or off.|
|&I|Turns italic printing on or off.|
|&U|Turns underline printing on or off.|
|&S|Turns strikethrough printing on or off.|
|&"fontname"|Prints the characters that follow in the specified font. Be sure to include the double quotation marks.|
|&nn|Prints the characters that follow in the specified font size. Use a two-digit number to specify a size in points.|
|&color|Prints the characters in the specified color. User supplies a hexadecimal color value.|
|&"+"|Prints the characters that follow in the **Heading** font of the current theme. Be sure to include the double quotation marks.|
|&"-"|Prints the characters that follow in the **Body** font of the current theme. Be sure to include the double quotation marks.|
|&K _xx_. _S_ _nnn_|Prints the characters that follow in the specified color from the current theme.<br/><br/>_xx_ is a two-digit number from 1 to 12 that specifies the theme color to use.<br/><br/>_S_ _nnn_ specifies the shade (tint) of that theme color. Specify _S_ as + to produce a lighter shade; specify _S_ as - to produce a darker shade.<br/><br/>_nnn_ is a three-digit whole number that specifies a percentage from 0 to 100.<br/><br/>If the values that specify the theme color or shade are not within the described limits, Excel will use the nearest valid value.|

<br/>

|**VBA code**|**Description**|
|:-----|:-----|
|&D|Prints the current date.|
|&T|Prints the current time.|
|&F|Prints the name of the document.|
|&A|Prints the name of the workbook tab.|
|&P|Prints the page number.|
|&P+number|Prints the page number plus the specified number.|
|&P-number|Prints the page number minus the specified number.|
|&&|Prints a single ampersand.|
|&N|Prints the total number of pages in the document. |
|&Z|Prints the file path.|
|&G|Inserts an image.|

## Example

The following code shows how formatting and VBA codes can be used to modify the header information and appearance.

```vb
Sub Date_Time() 
 ActiveSheet.PageSetup.CenterHeader = "&D &B&ITime:&I&B&T" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
