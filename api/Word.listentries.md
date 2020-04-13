---
title: ListEntries object (Word)
ms.prod: word
ms.assetid: cfd3c706-5b69-338f-b104-e12526b89f47
ms.date: 06/08/2017
localization_priority: Normal
---


# ListEntries object (Word)

A collection of  **[ListEntry](Word.ListEntry.md)** objects that represent all the items in a drop-down form field.


## Remarks

Use the **ListEntries** property to return the **ListEntries** collection. The following example displays the items that appear in the form field named "Drop1."


```vb
For Each le In _ 
 ActiveDocument.FormFields("Drop1").DropDown.ListEntries 
 MsgBox le.Name 
Next le
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


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]