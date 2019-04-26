---
title: Hyperlink.Range property (Excel)
keywords: vbaxl10.chm536075
f1_keywords:
- vbaxl10.chm536075
ms.prod: excel
api_name:
- Excel.Hyperlink.Range
ms.assetid: 0fdc49ba-fd3f-1125-fe3c-481828b7319e
ms.date: 04/26/2019
localization_priority: Normal
---


# Hyperlink.Range property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range that the specified hyperlink is attached to.


## Syntax

_expression_.**Range**

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


## Example

The following example stores in a variable the address for the AutoFilter applied to the Crew worksheet.

```vb
rAddress = Worksheets("Crew").AutoFilter.Range.Address
```

<br/>

This example scrolls through the workbook window until the hyperlink range is in the upper-left corner of the active window.

```vb
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).Range 
ActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]