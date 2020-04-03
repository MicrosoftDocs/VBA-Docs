---
title: FormFields.Add method (Word)
keywords: vbawd10.chm153682021
f1_keywords:
- vbawd10.chm153682021
ms.prod: word
api_name:
- Word.FormFields.Add
ms.assetid: d4431691-c881-e3b4-d17d-86c8ce07cf68
ms.date: 06/08/2017
localization_priority: Normal
---


# FormFields.Add method (Word)

Returns a  **FormField** object that represents a new form field added at a range.


## Syntax

_expression_.**Add** (_Range_, _Type_)

_expression_ Required. A variable that represents a '[FormFields](Word.formfields.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want to add the form field. If the range isn't collapsed, the form field replaces the range.|
| _Type_|Required| **WdFieldType**|The type of form field to add.|

## Return value

FormField


## Example

This example adds a check box at the end of the selection, gives it a name, and then selects it.


```vb
Selection.Collapse Direction:=wdCollapseEnd 
Set ffield = ActiveDocument.FormFields _ 
 .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox) 
With ffield 
 .Name = "Check_Box_1" 
 .CheckBox.Value = True 
End With
```


## See also


[FormFields Collection Object](Word.formfields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]