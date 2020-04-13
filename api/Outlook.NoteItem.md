---
title: NoteItem object (Outlook)
keywords: vbaol11.chm3001
f1_keywords:
- vbaol11.chm3001
ms.prod: outlook
api_name:
- Outlook.NoteItem
ms.assetid: ddf5baaa-6e13-a6fb-96e8-311e7761fa98
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem object (Outlook)

Represents a note in a Notes folder.


## Remarks

A **NoteItem** is not customizable. If you open a new note, you will notice that it is not possible to place it in design time.

The **[Subject](Outlook.NoteItem.Subject.md)** property of a **NoteItem** object is read-only because it is calculated from the body text of the note. Also, the **NoteItem** **[Body](Outlook.NoteItem.Body.md)** can only be rich text, so the properties that correspond to HTML and Microsoft Word content do not apply. Although the **[GetInspector](Outlook.NoteItem.GetInspector.md)** property will work on notes, because notes can't be customized, some of the **[Inspector](Outlook.Inspector.md)** properties, methods, and events will not apply to **NoteItem** objects.

Use the  **[CreateItem](Outlook.Application.CreateItem.md)** method to create a **NoteItem** object that represents a new note.

Use  **[Items](Outlook.Items.Item.md)** (_index_), where _index_ is the index number of a note or a value used to match the default property of a note, to return a single **NoteItem** object from a Notes folder.


## Example

 The following Microsoft Visual Basic example returns a new note.


```vb
Set myItem = Application.CreateItem(olNoteItem)
```


## Methods



|Name|
|:-----|
|[Close](Outlook.NoteItem.Close.md)|
|[Copy](Outlook.NoteItem.Copy.md)|
|[Delete](Outlook.NoteItem.Delete.md)|
|[Display](Outlook.NoteItem.Display.md)|
|[Move](Outlook.NoteItem.Move.md)|
|[PrintOut](Outlook.NoteItem.PrintOut.md)|
|[Save](Outlook.NoteItem.Save.md)|
|[SaveAs](Outlook.NoteItem.SaveAs.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.NoteItem.Application.md)|
|[AutoResolvedWinner](Outlook.NoteItem.AutoResolvedWinner.md)|
|[Body](Outlook.NoteItem.Body.md)|
|[Categories](Outlook.NoteItem.Categories.md)|
|[Class](Outlook.NoteItem.Class.md)|
|[Conflicts](Outlook.NoteItem.Conflicts.md)|
|[CreationTime](Outlook.NoteItem.CreationTime.md)|
|[DownloadState](Outlook.NoteItem.DownloadState.md)|
|[EntryID](Outlook.NoteItem.EntryID.md)|
|[GetInspector](Outlook.NoteItem.GetInspector.md)|
|[Height](Outlook.NoteItem.Height.md)|
|[IsConflict](Outlook.NoteItem.IsConflict.md)|
|[ItemProperties](Outlook.NoteItem.ItemProperties.md)|
|[LastModificationTime](Outlook.NoteItem.LastModificationTime.md)|
|[Left](Outlook.NoteItem.Left.md)|
|[MarkForDownload](Outlook.NoteItem.MarkForDownload.md)|
|[MessageClass](Outlook.NoteItem.MessageClass.md)|
|[Parent](Outlook.NoteItem.Parent.md)|
|[PropertyAccessor](Outlook.NoteItem.PropertyAccessor.md)|
|[Saved](Outlook.NoteItem.Saved.md)|
|[Session](Outlook.NoteItem.Session.md)|
|[Size](Outlook.NoteItem.Size.md)|
|[Subject](Outlook.NoteItem.Subject.md)|
|[Top](Outlook.NoteItem.Top.md)|
|[Width](Outlook.NoteItem.Width.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]