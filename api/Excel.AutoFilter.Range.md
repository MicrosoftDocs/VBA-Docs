---
title: AutoFilter.Range property (Excel)
keywords: vbaxl10.chm538073
f1_keywords:
- vbaxl10.chm538073
ms.prod: excel
api_name:
- Excel.AutoFilter.Range
ms.assetid: f8d1aca1-0d69-161a-981a-4dd10826e9d6
ms.date: 04/13/2019
localization_priority: Normal
---


# AutoFilter.Range property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range to which the specified **AutoFilter** applies.


## Syntax

_expression_.**Range**

_expression_ A variable that represents an **[AutoFilter](Excel.AutoFilter.md)** object.


## Example

The following example stores in a variable the address for the **AutoFilter** applied to the Crew worksheet.

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
