---
title: Phonetics object (Excel)
keywords: vbaxl10.chm657072
f1_keywords:
- vbaxl10.chm657072
ms.prod: excel
api_name:
- Excel.Phonetics
ms.assetid: 77c0c55c-a181-c68a-24ed-e6bcaf514663
ms.date: 06/08/2017
localization_priority: Normal
---


# Phonetics object (Excel)

A collection of all the  **[Phonetic](Excel.Phonetic.md)** objects in the specified range.


## Remarks

Each  **Phonetic** object contains information about a specific phonetic text string.


## Example

Use the  **[Phonetics](Excel.Range.Phonetics.md)** property to return the **Phonetics** collection. The following example makes all phonetic text in the range A1:C4 visible.


```vb
Range("A1:C4").Phonetics.Visible = True
```

Use  **Phonetics** ( _index_ ), where _index_ is the index number of the phonetic text, to return a single **Phonetic** object. The following example sets the first phonetic text string in the active cell to "フリガナ".




```vb
ActiveCell.Phonetics(1).Text = "フリガナ"
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


