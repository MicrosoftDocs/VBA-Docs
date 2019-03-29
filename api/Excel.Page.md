---
title: Page object (Excel)
keywords: vbaxl10.chm830072
f1_keywords:
- vbaxl10.chm830072
ms.prod: excel
api_name:
- Excel.Page
ms.assetid: debd4537-af71-8699-b714-6854c3cf0fad
ms.date: 03/30/2019
localization_priority: Normal
---


# Page object (Excel)

Represents a page in a workbook. Use the **Page** object and the related methods and properties for programmatically defining page layout in a workbook.


## Remarks

Use the **[Item](excel.pages.item.md)** property to access a specific page in a workbook. The following example accesses the first page in the active workbook.

```vb
Dim objPage As Page 
 
Set objPage = ActiveWorkbook.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```

## Properties

- [CenterFooter](Excel.Page.CenterFooter.md)
- [CenterHeader](Excel.Page.CenterHeader.md)
- [LeftFooter](Excel.Page.LeftFooter.md)
- [LeftHeader](Excel.Page.LeftHeader.md)
- [RightFooter](Excel.Page.RightFooter.md)
- [RightHeader](Excel.Page.RightHeader.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]