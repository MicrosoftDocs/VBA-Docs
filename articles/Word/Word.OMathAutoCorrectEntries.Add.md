---
title: OMathAutoCorrectEntries.Add Method (Word)
keywords: vbawd10.chm247988424
f1_keywords:
- vbawd10.chm247988424
ms.prod: word
api_name:
- Word.OMathAutoCorrectEntries.Add
ms.assetid: 0ef66b97-9da4-652d-306d-34e22945713c
ms.date: 06/08/2017
---


# OMathAutoCorrectEntries.Add Method (Word)

Creates an equation auto correct entry and returns an  **[OMathAutoCorrectEntry](Word.OMathAutoCorrectEntry.md)** object.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Value_** )

 _expression_ An expression that returns an **[OMathAutoCorrectEntries](Word.OMathAutoCorrectEntries.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the autocorrect entry. Corresponds to the  **[Name](Word.OMathAutoCorrectEntry.Name.md)** property of the **OMathAutoCorrectEntry** object.|
| _Value_|Required| **String**|The value of the autocorrect entry. Corresponds to the  **[Value](Word.OMathAutoCorrectEntry.Value.md)** property of the **OMathAutoCorrectEntry** object.|

### Return Value

OMathAutoCorrectEntry


## See also


#### Concepts


[OMathAutoCorrectEntries Collection](Word.OMathAutoCorrectEntries.md)

