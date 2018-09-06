---
title: HPageBreak Object (Excel)
keywords: vbaxl10.chm158072
f1_keywords:
- vbaxl10.chm158072
ms.prod: excel
api_name:
- Excel.HPageBreak
ms.assetid: 8fc96958-33ab-8251-f627-4769b5eab97f
ms.date: 06/08/2017
---


# HPageBreak Object (Excel)

Represents a horizontal page break. 


## Remarks

The  **HPageBreak** object is a member of the **[HPageBreaks](Excel.HPageBreaks.md)** collection.


 **Note**  There is a limit of 1026 horizontal page breaks per sheet.


## Example

Use  **[HPageBreaks](Excel.Worksheets.HPageBreaks.md)** ( _index_ ), where _index_ is the index number of the page break, to return an **HPageBreak** object. The following example changes the location of horizontal page break one.


```vb
Set Worksheets(1).HPageBreaks(1).Location = Worksheets(1).Range("e5")
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


