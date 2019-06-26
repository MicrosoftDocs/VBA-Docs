---
title: ListTemplate object (Word)
keywords: vbawd10.chm2447
f1_keywords:
- vbawd10.chm2447
ms.prod: word
api_name:
- Word.ListTemplate
ms.assetid: d5e339f7-5798-305b-a6b0-6b572d9112f4
ms.date: 06/08/2017
localization_priority: Normal
---


# ListTemplate object (Word)

Represents a single list template that includes all the formatting that defines a list. The  **ListTemplate** object is a member of the **ListTemplates** collection.


## Remarks

Each of the seven formats (excluding  **None**) found on each of the three tabs in the  **Bullets and Numbering** dialog box corresponds to a list template object. These predefined list templates can be accessed from the three **ListGallery** objects in the **ListGalleries** collection. Documents and templates can also contain collections of list templates.

Use  **ListTemplates** (Index), where Index is a number from 1 through 7, to return a single list template from a list gallery. The following example returns the third list format (excluding **None**) on the  **Numbered** tab in the **Bullets and Numbering** dialog box.




```vb
Set temp3 = ListGalleries(2).ListTemplates(3)
```


> [!NOTE] 
> Some properties and methods —  **Convert** and **Add**, for example — won't work with list templates that are accessed from a list gallery. You can modify these list templates, but you cannot change their list gallery type (**wdBulletGallery**, **wdNumberGallery**, or **wdOutlineNumberGallery**).

The following example sets an object variable equal to the list template used in the third list in the active document, and then it applies that list template to the selection.




```vb
Set myLt = ActiveDocument.ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myLt
```

Use the  **Add** method to add a list template to the collection of list templates in a document or template.

To see whether the specified list template contains the formatting built into Word, use the  **Modified** property with the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.

After you have returned a  **ListTemplate** object, use **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **ListLevel** object. With a **ListLevel** object, you have access to all the formatting properties for the specified list level, such as **Alignment**, **Font**, **NumberFormat**, **NumberPosition**, **NumberStyle**, and **TrailingCharacter**.

Use the  **Convert** method to convert a multiple-level list template to a single-level template.


## Methods



|Name|
|:-----|
|[Convert](Word.ListTemplate.Convert.md)|

## Properties



|Name|
|:-----|
|[Application](Word.ListTemplate.Application.md)|
|[Creator](Word.ListTemplate.Creator.md)|
|[ListLevels](Word.ListTemplate.ListLevels.md)|
|[Name](Word.ListTemplate.Name.md)|
|[OutlineNumbered](Word.ListTemplate.OutlineNumbered.md)|
|[Parent](Word.ListTemplate.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]