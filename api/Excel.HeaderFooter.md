---
title: HeaderFooter object (Excel)
keywords: vbaxl10.chm829072
f1_keywords:
- vbaxl10.chm829072
ms.prod: excel
api_name:
- Excel.HeaderFooter
ms.assetid: 75c654df-d3f9-8448-8a7e-a0487ca0d1ab
ms.date: 03/30/2019
localization_priority: Normal
---


# HeaderFooter object (Excel)

Represents a single header or footer. The **HeaderFooter** object is a member of the **HeadersFooters** collection.


## Remarks

You can also return a single **HeaderFooter** object by using the **HeaderFooter** property with a **Selection** object.

> [!NOTE] 
> You cannot add **HeaderFooter** objects to the **HeadersFooters** collection.

Use the **[DifferentFirstPageHeaderFooter](excel.pagesetup.differentfirstpageheaderfooter.md)** property of the **PageSetup** object to specify a different first page.


## Example

The following example adds the date and time to the center header on the active worksheet.

```vb
With ActiveSheet.PageSetup 
 .CenterHeader = "&D&T" 
 .OddAndEvenPagesHeaderFooter = False 
 .DifferentFirstPageHeaderFooter = False 
 .ScaleWithDocHeaderFooter = True 
 .AlignMarginsHeaderFooter = True 
End With
```

## Properties

- [Picture](Excel.HeaderFooter.Picture.md)
- [Text](Excel.HeaderFooter.Text.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]