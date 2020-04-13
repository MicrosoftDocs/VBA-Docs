---
title: ListGallery object (Word)
keywords: vbawd10.chm2452
f1_keywords:
- vbawd10.chm2452
ms.prod: word
api_name:
- Word.ListGallery
ms.assetid: 4fa3af33-becd-0dfc-5c7a-a0e70714e045
ms.date: 06/08/2017
localization_priority: Normal
---


# ListGallery object (Word)

Represents a single gallery of list formats. The **ListGallery** object is a member of the **ListGalleries** collection.


## Remarks

Each  **ListGallery** object represents one of the three tabs in the **Bullets and Numbering** dialog box.

Use  **ListGalleries** (Index), where Index is **wdBulletGallery**, **wdNumberGallery**, or **wdOutlineNumberGallery**, to return a single **ListGallery** object.

The following example returns the third list format (excluding  **None**) on the **Bulleted** tab in the **Bullets and Numbering** dialog box and then applies it to the selection.




```vb
Set temp3 = ListGalleries(wdBulletGallery).ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:= temp3
```

To see whether the specified list template contains the formatting built into Word, use the **Modified** property for the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.


## Methods



|Name|
|:-----|
|[Reset](Word.ListGallery.Reset.md)|

## Properties



|Name|
|:-----|
|[Application](Word.ListGallery.Application.md)|
|[Creator](Word.ListGallery.Creator.md)|
|[ListTemplates](Word.ListGallery.ListTemplates.md)|
|[Modified](Word.ListGallery.Modified.md)|
|[Parent](Word.ListGallery.Parent.md)|


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]