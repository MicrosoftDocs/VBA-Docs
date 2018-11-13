---
title: VPageBreak object (Excel)
keywords: vbaxl10.chm155072
f1_keywords:
- vbaxl10.chm155072
ms.prod: excel
api_name:
- Excel.VPageBreak
ms.assetid: 0b37bdc0-b7e2-2b3f-ba6c-853cbbb67837
ms.date: 06/08/2017
---


# VPageBreak object (Excel)

Represents a vertical page break.


## Remarks

The  **VPageBreak** object is a member of the **[VPageBreaks](Excel.VPageBreaks.md)** collection.


## Example

Use  **VPageBreaks** ( _index_), where  _index_ is the page break index number of the page break, to return a **VPageBreak** object. The following example changes the location of vertical page break one.


```vb
Worksheets(1).VPageBreaks(1).Location = Worksheets(1).Range("e5")
```


## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)


