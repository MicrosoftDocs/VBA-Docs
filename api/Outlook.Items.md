---
title: Items object (Outlook)
keywords: vbaol11.chm2998
f1_keywords:
- vbaol11.chm2998
ms.prod: outlook
api_name:
- Outlook.Items
ms.assetid: 3a99730b-e62a-5ca6-f6ec-911c95173242
ms.date: 06/08/2017
localization_priority: Normal
---


# Items object (Outlook)

Contains a collection of [Outlook item objects](../outlook/How-to/Items-Folders-and-Stores/outlook-item-objects.md) in a folder.


## Remarks

Use the **[Items](Outlook.Folder.Items.md)** property to return the **Items** object of a **[Folder](Outlook.Folder.md)** object.

Use **Items** (_index_), where _index_ is the name or index number, to return a single Outlook item.


> [!NOTE] 
> The index for the **Items** collection starts at 1, and the items in the **Items** collection object are not guaranteed to be in any particular order.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns the first item in the **Inbox** with the Subject "Need your advice."






```vb
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items("Need your advice") 
 
 myItem.Display 
 
End sub
```

The following VBA example returns the first item in the **Inbox**. In Microsoft Office Outlook 2003 or later, the **Items** object returns the items in an Offline Folders file (.ost) in the reverse order.






```vb
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items(1) 
 
 myItem.Display 
 
End sub
```


## Events



|Name|
|:-----|
|[ItemAdd](Outlook.Items.ItemAdd.md)|
|[ItemChange](Outlook.Items.ItemChange.md)|
|[ItemRemove](Outlook.Items.ItemRemove.md)|

## Methods



|Name|
|:-----|
|[Add](Outlook.Items.Add.md)|
|[Find](Outlook.Items.Find.md)|
|[FindNext](Outlook.Items.FindNext.md)|
|[GetFirst](Outlook.Items.GetFirst.md)|
|[GetLast](Outlook.Items.GetLast.md)|
|[GetNext](Outlook.Items.GetNext.md)|
|[GetPrevious](Outlook.Items.GetPrevious.md)|
|[Item](Outlook.Items.Item.md)|
|[Remove](Outlook.Items.Remove.md)|
|[ResetColumns](Outlook.Items.ResetColumns.md)|
|[Restrict](Outlook.Items.Restrict.md)|
|[SetColumns](Outlook.Items.SetColumns.md)|
|[Sort](Outlook.Items.Sort.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Items.Application.md)|
|[Class](Outlook.Items.Class.md)|
|[Count](Outlook.Items.Count.md)|
|[IncludeRecurrences](Outlook.Items.IncludeRecurrences.md)|
|[Parent](Outlook.Items.Parent.md)|
|[Session](Outlook.Items.Session.md)|

## See also


[Items Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
