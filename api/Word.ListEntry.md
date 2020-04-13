---
title: ListEntry object (Word)
keywords: vbawd10.chm2339
f1_keywords:
- vbawd10.chm2339
ms.prod: word
api_name:
- Word.ListEntry
ms.assetid: ea9e8276-45d6-8b11-fd86-4944f582bb80
ms.date: 06/08/2017
localization_priority: Normal
---


# ListEntry object (Word)

Represents an item in a drop-down form field. The **ListEntry** object is a member of the **ListEntries** collection. The **[ListEntries](Word.listentries.md)** collection includes all the items in a drop-down form field.


## Remarks

Use  **ListEntries** (Index), where Index is the list entry name or the index number, to return a single **ListEntry** object. The index number represents the position of the entry in the drop-down form field (the first item is index number 1). The following example deletes the "Blue" entry from the drop-down form field named "Color."


```vb
ActiveDocument.FormFields("Color").DropDown _ 
 .ListEntries("Blue").Delete
```

The following example displays the first item in the drop-down form field named "Color."




```vb
MsgBox _ 
 ActiveDocument.FormFields("Color").DropDown.ListEntries(1).Name
```

Use the **Add** method to add an item to a drop-down form field. The following example inserts a drop-down form field and then adds "red," "blue," and "green" to the form field.




```vb
Set myField = _ 
 ActiveDocument.FormFields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldFormDropDown) 
With myField.DropDown.ListEntries 
 .Add Name:="Red" 
 .Add Name:="Blue" 
 .Add Name:="Green" 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]