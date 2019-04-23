---
title: CheckBox object (Word)
keywords: vbawd10.chm2342
f1_keywords:
- vbawd10.chm2342
ms.prod: word
api_name:
- Word.CheckBox
ms.assetid: e72b57b7-0328-9e78-94ca-ab7fb3c64afb
ms.date: 06/08/2017
localization_priority: Normal
---


# CheckBox object (Word)

Represents a single check box form field.


## Remarks

Use  **FormFields** (Index), where Index is index number or the bookmark name associated with the check box, to return a single **[FormField](Word.FormField.md)** object. Use the **[CheckBox](Word.FormField.CheckBox.md)** property with the **FormField** object to return a **CheckBox** object. The following example selects the check box form field named "Check1" in the active document.


```vb
ActiveDocument.FormFields("Check1").CheckBox.Value = True
```

The index number represents the position of the form field in the  **[FormFields](Word.formfields.md)** collection. The following example checks the type of the first form field; if it is a check box, the check box is selected.




```vb
If ActiveDocument.FormFields(1).Type = wdFieldFormCheckBox Then 
 ActiveDocument.FormFields(1).CheckBox.Value = True 
End If
```

The following example determines whether the  _ffield_ object is valid before changing the check box size to 14 points.




```vb
Set ffield = ActiveDocument.FormFields(1).CheckBox 
If ffield.Valid = True Then 
 ffield.AutoSize = False 
 ffield.Size = 14 
Else 
 MsgBox "First field is not a check box" 
End If
```

Use the  **Add** method with the **FormFields** object to add a check box form field. The following example adds a check box at the beginning of the active document, sets the name to "Color", and then selects the check box.




```vb
With ActiveDocument.FormFields.Add(Range:=ActiveDocument.Range _ 
 (Start:=0,End:=0), Type:=wdFieldFormCheckBox) 
 .Name = "Color" 
 .CheckBox.Value = True 
End With
```


## Properties



|Name|
|:-----|
|[Application](Word.CheckBox.Application.md)|
|[AutoSize](Word.CheckBox.AutoSize.md)|
|[Creator](Word.CheckBox.Creator.md)|
|[Default](Word.CheckBox.Default.md)|
|[Parent](Word.CheckBox.Parent.md)|
|[Size](Word.CheckBox.Size.md)|
|[Valid](Word.CheckBox.Valid.md)|
|[Value](Word.CheckBox.Value.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
