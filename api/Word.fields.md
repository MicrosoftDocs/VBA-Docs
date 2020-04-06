---
title: Fields object (Word)
ms.prod: word
ms.assetid: c79065bb-ba29-22fd-a9d7-90bb10550035
ms.date: 06/08/2017
localization_priority: Normal
---


# Fields object (Word)

A collection of  **Field** objects that represent all the fields in a selection, range, or document.


## Remarks

Use the  **Fields** property to return the **Fields** collection. The following example updates all the fields in the selection.


> [!NOTE] 
> Use the  **Fields** property with a **[MailMerge](Word.MailMerge.md)** object to return a **[MailMergeFields](Word.mailmergefields.md)** collection.


```vb
Selection.Fields.Update
```

Use the  **Add** method to add a field to the **Fields** collection. The following example inserts a DATE field at the beginning of the selection and then displays the result.




```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myField = ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldDate) 
MsgBox myField.Result
```

Use  **Fields** (Index), where Index is the index number, to return a single **[Field](Word.Field.md)** object. The index number represents the position of the field in the selection, range, or document. The following example displays the field code and the result of the first field in the active document.




```vb
If ActiveDocument.Fields.Count >= 1 Then 
 MsgBox "Code = " & ActiveDocument.Fields(1).Code & vbCr _ 
 & "Result = " & ActiveDocument.Fields(1).Result & vbCr 
End If
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
