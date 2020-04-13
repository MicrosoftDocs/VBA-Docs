---
title: AutoTextEntry.Insert method (Word)
keywords: vbawd10.chm154533990
f1_keywords:
- vbawd10.chm154533990
ms.prod: word
api_name:
- Word.AutoTextEntry.Insert
ms.assetid: 381e69fa-10c8-5951-e890-3fe8c508e047
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoTextEntry.Insert method (Word)

Inserts the AutoText entry in place of the specified range. Returns a  **Range** object that represents the AutoText entry.


## Syntax

_expression_.**Insert** (_Where_, _RichText_)

_expression_ Required. A variable that represents an '[AutoTextEntry](Word.AutoTextEntry.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Where_|Required| **Range**|The location for the AutoText entry.|
| _RichText_|Optional| **Variant**| **True** to insert the AutoText entry with its original formatting.|

## Return value

Range


## Remarks

If you don't want to replace the range, use the **Collapse** method before using this method.


## Example

This example inserts the formatted AutoText entry named "one" after the selection.


```vb
Sub InsertAutoTextEntry() 
 ActiveDocument.Content.Select 
 Selection.Collapse Direction:=wdCollapseEnd 
 ActiveDocument.AttachedTemplate.AutoTextEntries("one").Insert _ 
 Where:=Selection.Range, RichText:=True 
End Sub
```


## See also


[AutoTextEntry Object](Word.AutoTextEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]