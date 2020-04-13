---
title: DropDown object (Word)
keywords: vbawd10.chm2341
f1_keywords:
- vbawd10.chm2341
ms.prod: word
api_name:
- Word.DropDown
ms.assetid: 55233d61-d6d0-30f9-6825-ebbdbeb928b6
ms.date: 06/08/2017
localization_priority: Normal
---


# DropDown object (Word)

Represents a drop-down form field that contains a list of items in a form.


## Remarks

Use  **FormFields** (_index_), where _index_ is the index number or the bookmark name associated with the drop-down form field, to return a single **FormField** object. Use the **DropDown** property with the **FormField** object to return a **DropDown** object. The following example selects the first item in the drop-down form field named "DropDown" in the active document.


```vb
ActiveDocument.FormFields("DropDown1").DropDown.Value = 1
```

The index number represents the position of the form field in the **[FormFields](Word.formfields.md)** collection. The following example checks the type of the first form field in the active document. If it is a drop-down form field, the second item is selected.




```vb
If ActiveDocument.FormFields(1).Type = wdFieldFormDropDown Then 
 ActiveDocument.FormFields(1).DropDown.Value = 2 
End If
```

The following example determines whether form field represented by  _ffield_ is a valid drop-down form field before adding an item to it.




```vb
Set ffield = ActiveDocument.FormFields(1).DropDown 
If ffield.Valid = True Then 
 ffield.ListEntries.Add Name:="Hello" 
Else 
 MsgBox "First field is not a drop down" 
End If
```

Use the **Add** method with the **FormFields** collection to add a drop-down form field. The following example adds a drop-down form field at the beginning of the active document and then adds items to the form field.




```vb
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0, End:=0), _ 
 Type:=wdFieldFormDropDown) 
With ffield 
 .Name = "Colors" 
 With .DropDown.ListEntries 
 .Add Name:="Blue" 
 .Add Name:="Green" 
 .Add Name:="Red" 
 End With 
End With
```


## Properties



|Name|
|:-----|
|[Application](Word.DropDown.Application.md)|
|[Creator](Word.DropDown.Creator.md)|
|[Default](Word.DropDown.Default.md)|
|[ListEntries](Word.DropDown.ListEntries.md)|
|[Parent](Word.DropDown.Parent.md)|
|[Valid](Word.DropDown.Valid.md)|
|[Value](Word.DropDown.Value.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
