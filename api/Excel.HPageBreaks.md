---
title: HPageBreaks Object (Excel)
keywords: vbaxl10.chm163072
f1_keywords:
- vbaxl10.chm163072
ms.prod: excel
api_name:
- Excel.HPageBreaks
ms.assetid: 087106a7-ded7-d672-095d-98e7012fa440
ms.date: 06/08/2017
---


# HPageBreaks Object (Excel)

The collection of horizontal page breaks within the print area.


## Remarks

 Each horizontal page break is represented by an **[HPageBreak](Excel.HPageBreak.md)** object.

If you add a page break that does not intersect the print area, then the newly-added  **HPageBreak** object will not appear in the **HPageBreaks** collection for the print area. The contents of the collection may change if the print area is resized or redefined.

When the  **[Application](Excel.HPageBreaks.Application.md)** property, **[Count](Excel.HPageBreaks.Count.md)** property, **[Item](Excel.HPageBreaks.Item.md)** property, **[Parent](Excel.HPageBreaks.Parent.md)** property or **[Add](Excel.HPageBreaks.Add.md)** method is used in conjunction with the **HPageBreaks** property:


- For an automatic print area, the  **[HPageBreaks](Excel.Worksheet.HPageBreaks.md)** property applies only to the page breaks within the print area.
    
- For a user-defined print area of the same range, the  **HPageBreaks** property applies to all of the page breaks.
    

 **Note**  There is a limit of 1026 horizontal page breaks per sheet.


## Example

Use the  **HPageBreaks** property to return the **HPageBreaks** collection. Use the **Add** method to add a horizontal page break. The following example adds a horizontal page break above the active cell.


```vb
ActiveSheet.HPageBreaks.Add Before:=ActiveCell
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

