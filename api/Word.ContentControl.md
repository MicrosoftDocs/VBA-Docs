---
title: ContentControl object (Word)
keywords: vbawd10.chm4067
f1_keywords:
- vbawd10.chm4067
ms.prod: word
api_name:
- Word.ContentControl
ms.assetid: 783dec26-9b63-11f8-6187-985f9c815f27
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControl object (Word)

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. The **ContentControl** object is a member of the **[ContentControls](Word.ContentControls.md)** collection.


## Remarks

Use the **[Add](Word.ContentControls.Add.md)** method of the **ContentControls** collection to create a content control. Use the Type parameter of the **Add** method to specify the type of content control to create. The following example create a new drop-down list content control and adds several items to the list.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(Type:=wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```

Use the **Type** property to change the content control to a different type of content control. For example, perhaps you want to change from a date control to a text control. However, you may not be able to change all content controls to another type; some may not allow changing their type. In addition, depending on the contents of a content control, you may not be able to change the type. For example, if the content control that you want to change to does not allow the type of content that is in the existing content control, attempting to change the type is not allowed and generates a run-time error.

The following example inserts a date content control and sets the value of the control, and then changes the control to a text content control.




```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDate) 
 
objCC.Range.Text = "January 1, 2007" 
objCC.Type = wdContentControlText
```

Use the **SetPlaceholderText** method to change the placeholder text from the default string to something more appropriate for the control. Use the **Title** property to specify the title text for the control. This displays above the control when the cursor is positioned inside the control or the mouse pointer is positioned over the control.

Depending on the type of content control that you have, you may not be able to use all the properties and methods of the **ContentControl** object.

Not all content control properties apply to all the different types of content controls. The following table lists which properties apply to which types of content controls.



|**Property/Method**|**Applies To**|
|:-----|:-----|
| **[BuildingBlockCategory](Word.ContentControl.BuildingBlockCategory.md)** property|BuildingBlock Gallery content controls (wdContentControlBuildingBlockGallery)|
| **[BuildingBlockType](Word.ContentControl.BuildingBlockType.md)** property|BuildingBlock Gallery content controls (wdContentControlBuildingBlockGallery)|
| **[DateDisplayFormat](Word.ContentControl.DateDisplayFormat.md)** property|Date content controls (wdContentControlDate)|
| **[DateDisplayLocale](Word.ContentControl.DateDisplayLocale.md)** property|Date content controls (wdContentControlDate)|
| **[DateStorageFormat](Word.ContentControl.DateStorageFormat.md)** property|Date content controls (wdContentControlDate)|
| **[DropdownListEntries](Word.ContentControl.DropdownListEntries.md)** property|Combo box and drop-down list content controls (wdContentControlComboBox and wdContentControlDropdownList)|
| **[MultiLine](Word.ContentControl.MultiLine.md)** property|Plain text content controls (wdContentControlText)|
| **[Ungroup](Word.ContentControl.Ungroup.md)** method|Group content controls (wdContentControlGroup)|

## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
