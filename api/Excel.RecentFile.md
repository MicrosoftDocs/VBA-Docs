---
title: RecentFile Object (Excel)
keywords: vbaxl10.chm169072
f1_keywords:
- vbaxl10.chm169072
ms.prod: excel
api_name:
- Excel.RecentFile
ms.assetid: 39d0a969-179d-a7bd-e5ab-7baf7930712a
ms.date: 06/08/2017
---


# RecentFile Object (Excel)

Represents a file in the list of recently used files.


## Remarks

 The **RecentFile** object is a member of the **[RecentFiles](Excel.RecentFiles.md)** collection.


## Example

Use  **[RecentFiles](Excel.Application.RecentFiles.md)** ( _index_ ), where _index_ is the file number, to return a **RecentFile** object. The following example opens file two in the list of recently used files.


```vb
Application.RecentFiles(2).Open
```


## See also


[Excel Object Model Reference](./overview/object-model-excel-vba-reference.md)


