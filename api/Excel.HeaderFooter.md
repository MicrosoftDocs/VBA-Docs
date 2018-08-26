---
title: HeaderFooter Object (Excel)
keywords: vbaxl10.chm829072
f1_keywords:
- vbaxl10.chm829072
ms.prod: excel
api_name:
- Excel.HeaderFooter
ms.assetid: 75c654df-d3f9-8448-8a7e-a0487ca0d1ab
ms.date: 06/08/2017
---


# HeaderFooter Object (Excel)

Represents a single header or footer. The  **HeaderFooter** object is a member of the **HeadersFooters** collection.


## Remarks

You can also return a single  **HeaderFooter** object by using the **HeaderFooter** property with a **Selection** object.


 **Note**  You cannot add  **HeaderFooter** objects to the **HeadersFooters** collection.

Use the  **DifferentFirstPageHeaderFooter** property with the **PageSetup** object to specify a different first page.


## Example

The following example adds the date and time to the center header in the active worksheet.


```vb
With ActiveSheet.PageSetup 
<<<<<<< HEAD
 .CenterHeader = "&;D&;T" 
=======
 .CenterHeader = "&D&T" 
>>>>>>> master
 .OddAndEvenPagesHeaderFooter = False 
 .DifferentFirstPageHeaderFooter = False 
 .ScaleWithDocHeaderFooter = True 
 .AlignMarginsHeaderFooter = True 
End With
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

