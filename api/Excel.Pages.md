---
title: Pages object (Excel)
keywords: vbaxl10.chm831072
f1_keywords:
- vbaxl10.chm831072
ms.prod: excel
api_name:
- Excel.Pages
ms.assetid: ecedccc4-e1af-6a66-9d83-bd0cf76dfe68
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages object (Excel)

A collection of pages in a document. Use the  **Pages** collection and the related objects and properties for programmatically defining page layout in a workbook.


## Remarks

Use the  **Pages** property to return a **Pages** collection. The following example accesses all pages in the active worksheet.


```vb
Dim objPages As Pages 
 
Set objPage = ActiveWorksheet. _ 
 ActiveWindow.Panes(1).Pages
```

Use the  **Item** method to access an individual **Page** object that represents an individual page in a worksheet. The following example accesses the first page in the active worksheet.




```vb
Dim objPage As Page 
 
Set objPage = ActiveWorksheet.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

