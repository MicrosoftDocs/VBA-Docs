---
title: FormFields object (Word)
ms.prod: word
ms.assetid: a44a0f57-123b-cade-e306-ba6dc179b619
ms.date: 06/08/2017
localization_priority: Normal
---


# FormFields object (Word)

A collection of  **FormField** objects that represent all the form fields in a selection, range, or document.


## Remarks

Use the **FormFields** property to return the **FormFields** collection. The following example counts the number of text box form fields in the active document.


```vb
For Each aField In ActiveDocument.FormFields 
 If aField.Type = wdFieldFormTextInput Then count = count + 1 
Next aField 
MsgBox "There are " & count & " text boxes in this document"
```

Use the **Add** method with the **FormFields** object to add a form field. The following example adds a check box at the beginning of the active document and then selects the check box.




```vb
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0,End:=0), _ 
 Type:=wdFieldFormCheckBox) 
ffield.CheckBox.Value = True
```

Use  **FormFields** (Index), where Index is a bookmark name or index number, to return a single **[FormField](Word.FormField.md)** object. The following example sets the result of the Text1 form field to "Don Funk."




```vb
ActiveDocument.FormFields("Text1").Result = "Don Funk"
```

The index number represents the position of the form field in the selection, range, or document. The following example displays the name of the first form field in the selection.




```vb
If Selection.FormFields.Count >= 1 Then 
 MsgBox Selection.FormFields(1).Name 
End If
```


## Methods



|Name|
|:-----|
|[Add](Word.FormFields.Add.md)|
|[Item](Word.FormFields.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.FormFields.Application.md)|
|[Count](Word.FormFields.Count.md)|
|[Creator](Word.FormFields.Creator.md)|
|[Parent](Word.FormFields.Parent.md)|
|[Shaded](Word.FormFields.Shaded.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]