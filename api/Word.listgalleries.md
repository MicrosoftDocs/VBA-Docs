---
title: ListGalleries object (Word)
ms.prod: word
ms.assetid: 3ae91fbf-fb7c-e96f-fd13-e4e4e9c4f09e
ms.date: 06/08/2017
localization_priority: Normal
---


# ListGalleries object (Word)

A collection of  **[ListGallery](Word.ListGallery.md)** objects that represent the three tabs in the **Bullets and Numbering** dialog box.


## Remarks

Use the  **ListGalleries** property to return the **ListGalleries** collection. The following code example enumerates the collection of list galleries and sets each of the seven list templates (formats) back to the list template format built into Word.


```vb
For Each lg In ListGalleries 
 For x = 1 To 7 
 lg.Reset(x) 
 Next x 
Next lg
```

Use  **ListGalleries** (Index), where Index is **wdBulletGallery**, **wdNumberGallery**, or **wdOutlineNumberGallery**, to return a single **ListGallery** object.

The following code example returns the third list format (excluding  **None**) on the  **Bulleted** tab in the **Bullets and Numbering** dialog box and then applies it to the selection.




```vb
Set temp3 = ListGalleries(wdBulletGallery).ListTemplates(3) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:= temp3
```

To see whether the specified list template contains the formatting built into Word, use the  **Modified** property with the **ListGallery** object. To reset formatting to the original list format, use the **Reset** method for the **ListGallery** object.


## Methods



|Name|
|:-----|
|[Item](Word.ListGalleries.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.ListGalleries.Application.md)|
|[Count](Word.ListGalleries.Count.md)|
|[Creator](Word.ListGalleries.Creator.md)|
|[Parent](Word.ListGalleries.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]