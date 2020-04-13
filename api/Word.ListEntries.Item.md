---
title: ListEntries.Item method (Word)
keywords: vbawd10.chm153354240
f1_keywords:
- vbawd10.chm153354240
ms.prod: word
api_name:
- Word.ListEntries.Item
ms.assetid: 749a78cf-b72e-defe-396b-cd7f3c802277
ms.date: 06/08/2017
localization_priority: Normal
---


# ListEntries.Item method (Word)

Returns an individual  **ListEntry** object in a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ Required. A variable that represents a '[ListEntries](Word.listentries.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The individual object to be returned. Can be a  **Long** indicating the ordinal position or a **String** representing the name of the individual object.|

## Return value

ListEntry


## Example

This example clears all the items from the drop-down form field named "Colors" and then adds two color names. The **Item** method is used to display the first color in the drop-down form field.


```vb
Sub ListEntryItem() 
 Dim d As DropDown 
 Set d = ActiveDocument.FormFields.Add _ 
 (Range:=Selection.Range, _ 
 Type:=wdFieldFormDropDown).DropDown 
 With d.ListEntries 
 .Add Name:="Black" 
 .Add Name:="Green" 
 End With 
 MsgBox d.ListEntries.Item(1).Name 
End Sub
```


## See also


[ListEntries Collection Object](Word.listentries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]