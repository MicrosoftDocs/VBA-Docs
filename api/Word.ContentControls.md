---
title: ContentControls object (Word)
ms.prod: word
api_name:
- Word.ContentControls
ms.assetid: 2595eea9-df68-edce-3a51-069cad14bb87
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControls object (Word)

A collection of  **[ContentControl](Word.ContentControl.md)** objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain content such as dates, lists, or paragraphs of formatted text.


## Remarks

Use the **[Add](Word.ContentControls.Add.md)** method to create a new content control and insert it into a document. The following example creates a new drop-down list content control and adds several items to the list.


```vb
Dim objcc As ContentControl 
Dim objMap As XMLMapping 
 
Set objcc = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
objcc.Title = "My Favorite Animal" 
If objcc.ShowingPlaceholderText Then _ 
 objcc.SetPlaceholderText , , "Select your favorite animal " 
 
'List entries 
objcc.DropdownListEntries.Add "Cat" 
objcc.DropdownListEntries.Add "Dog" 
objcc.DropdownListEntries.Add "Horse" 
objcc.DropdownListEntries.Add "Monkey" 
objcc.DropdownListEntries.Add "Snake" 
objcc.DropdownListEntries.Add "Other"
```

Use the **[Item](Word.ContentControls.Item.md)** method to access a specific content control in the collection. The following example accesses the third content control in the active document, and if the control is a drop-down list or a combo box, moves the first item to the bottom of the list and the last item to the top of the list.




```vb
Dim objcc As ContentControl 
Dim objLE1 As ContentControlListEntry 
Dim objLE2 As ContentControlListEntry 
Dim intCount As Integer 
 
Set objcc = ActiveDocument.ContentControls.Item(3) 
 
If objcc.Type = wdContentControlComboBox Or _ 
 objcc.Type = wdContentControlDropdownList Then 
 
 'First item in the list. 
 Set objLE1 = objcc.DropdownListEntries.Item(1) 
 
 'Last item in the list. 
 Set objLE2 = objcc.DropdownListEntries.Item(objcc.DropdownListEntries.Count) 
 
 For intCount = 1 To objcc.DropdownListEntries.Count 
 'Move the first item down one. 
 objLE1.MoveDown 
 
 'Move the last item up one. 
 objLE2.MoveUp 
 Next 
 
End If
```

Use the **ContentControl** object to work with individual content controls. For more information, see [Working with Content Controls](../word/Concepts/Working-with-Word/working-with-content-controls.md).


## Methods



|Name|
|:-----|
|[Add](Word.ContentControls.Add.md)|
|[Item](Word.ContentControls.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.ContentControls.Application.md)|
|[Count](Word.ContentControls.Count.md)|
|[Creator](Word.ContentControls.Creator.md)|
|[Parent](Word.ContentControls.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
