---
title: Range.InsertAlignmentTab Method (Word)
keywords: vbawd10.chm157155828
f1_keywords:
- vbawd10.chm157155828
ms.prod: word
api_name:
- Word.Range.InsertAlignmentTab
ms.assetid: 1ca21f95-ca53-e911-c789-b0203d7bf0c7
ms.date: 06/08/2017
---


# Range.InsertAlignmentTab Method (Word)

Inserts an absolute tab that is always positioned in the same spot, relative to either the margins or indents.


## Syntax

 _expression_ . **InsertAlignmentTab**( **_Alignment_** , **_RelativeTo_** )

 _expression_ An expression that returns a **[Range](Word.Range.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Alignment_|Required| **Long**|Indicates the type of alignment?left, center, or right?for the tab stop. Can be one of the  **[WdAlignmentTabAlignment](Word.WdAlignmentTabAlignment.md)** constants.|
| _RelativeTo_|Optional| **Long**|Indicates whether the tab stop is relative to the margins or to the paragraph indents. Can be one of the  **[WdAlignmentTabRelative](Word.WdAlignmentTabRelative.md)** constants.|

## Example

The following example inserts an alignment tab at the Insertion Point.


```
Selection.Range.InsertAlignmentTab 1, 1
```


## See also


#### Concepts


[Range Object](Word.Range.md)

