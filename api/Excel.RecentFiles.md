---
title: RecentFiles object (Excel)
keywords: vbaxl10.chm171072
f1_keywords:
- vbaxl10.chm171072
ms.prod: excel
api_name:
- Excel.RecentFiles
ms.assetid: e33ae942-0444-0631-be08-386366b6ebdb
ms.date: 06/08/2017
localization_priority: Normal
---


# RecentFiles object (Excel)

Represents the list of recently used files.


## Remarks

 Each file is represented by a **[RecentFile](Excel.RecentFile.md)** object.


## Example

Use the  **[RecentFiles](Excel.Application.RecentFiles.md)** property to return the **RecentFiles** collection. The following example sets the maximum number of files in the list of recently used files.


```vb
Application.RecentFiles.Maximum = 6
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


