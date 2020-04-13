---
title: ListTemplates object (Word)
ms.prod: word
ms.assetid: 5b5f3ed8-4522-f52e-5ae8-9df26a7da154
ms.date: 06/08/2017
localization_priority: Normal
---


# ListTemplates object (Word)

A collection of  **[ListTemplate](Word.listTemplate.md)** objects in a document, list gallery, or template.


## Remarks

Use the **ListTemplates** property with a [Document](Word.Document.md), [ListGallery](Word.ListGallery.md), or [Template](Word.Template.md) object to return a **ListTemplates** collection. With a ListGallery object, the ListTemplates collection is the seven list formats for bulleted lists, numbered lists, and outline numbered lists. 
 
The following example displays a message with the level status (single or multiple-level) for each list template in the active document.


```vb
For Each lt In ActiveDocument.ListTemplates 
 MsgBox "This is a multiple-level list template - " _ 
 & lt.OutlineNumbered 
Next lt
```

Use the **Add** method to add a list template to the collection in the specified document or template. The following example adds a new list template to the active document and applies it to the selection.




```vb
Set myLT = ActiveDocument.ListTemplates.Add 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myLT
```

Use  **ListTemplates** (Index), where Index is the name of a list template or an index number, to return a single list template in a document or template. The following example sets an object variable equal to a list template named "ListBullets" in the active document, and then formats the selection as the first level of that list template. 


```vb
Set mylt = ActiveDocument.ListTemplates("ListBullets")
Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:=mylt, ApplyLevel:=1
```

Use  **ListTemplates** (Index), where Index is a number 1 through 7, to return a single list template in a list gallery. The following example sets an object variable equal to the first list template in the bullet list gallery, and then it applies that list template to the selection.




```vb
Set mylt = ListGalleries(wdBulletGallery).ListTemplates(1) 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=mylt
```


>> [!NOTE] 
> Some properties and methods —  **Convert** and **Add**, for example — won't work with the list templates in a list gallery. You can modify those list templates, but you cannot change their list gallery type (**wdBulletGallery**, **wdNumberGallery**, or **wdOutlineNumberGallery**).

To see whether a list template in a list gallery contains the formatting built into Word, use the **[Modified](Word.ListGallery.Modified.md)** property with the **ListGallery** object. To reset formatting to the original list format, use the **[Reset](Word.ListGallery.Reset.md)** method for the **ListGallery** object.

After you have returned a  **[ListTemplate](Word.listTemplate.md)** object, use **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **ListLevel** object. With a **ListLevel** object, you have access to all the formatting properties for the specified list level, such as **Alignment**, **Font**, **NumberFormat**, **NumberPosition**, **NumberStyle**, and **TrailingCharacter**.

Use the **Convert** method to convert a multiple-level list template to a single-level template.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]