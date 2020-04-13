---
title: BusinessCardView object (Outlook)
keywords: vbaol11.chm3212
f1_keywords:
- vbaol11.chm3212
ms.prod: outlook
api_name:
- Outlook.BusinessCardView
ms.assetid: 83706cf8-080c-fbf0-9381-5801a2dd4dfd
ms.date: 06/08/2017
localization_priority: Normal
---


# BusinessCardView object (Outlook)

Represents a view that displays data as a series of Electronic Business Card (EBC) images.


## Remarks

The **BusinessCardView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to better sort, group and ultimately view contact items in Outlook as a series of Electronic Business Cards, each of which displays the contact information for an Outlook contact item based on the EBC design associated with the contact item.

Use the  **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **BusinessCardView** to a **[Folder](Outlook.Folder.md)** object.

Use the  **[Filter](Outlook.BusinessCardView.Filter.md)** property to determine which Outlook contact items to display in the view, the **[CardSize](Outlook.BusinessCardView.CardSize.md)** property to specify the size of each Electronic Business Card in the view, and the **[HeadingsFont](Outlook.BusinessCardView.HeadingsFont.md)** to retrieve the **[ViewFont](Outlook.ViewFont.md)** object for the view. Use the **[LockUserChanges](Outlook.BusinessCardView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.


## Example

The following Visual Basic for Applications (VBA) example creates, saves, and applies a new **BusinessCardView** object.


```vb
Sub CreateBusinessCardView() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Inbox default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Create the new view. 
 
 Set objView = objViews.Add( _ 
 
 "Card View", _ 
 
 olBusinessCardView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 
 
 ' Save and apply the new view. 
 
 objView.Save 
 
 objView.Apply 
 
 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Apply](Outlook.BusinessCardView.Apply.md)|
|[Copy](Outlook.BusinessCardView.Copy.md)|
|[Delete](Outlook.BusinessCardView.Delete.md)|
|[GoToDate](Outlook.BusinessCardView.GoToDate.md)|
|[Reset](Outlook.BusinessCardView.Reset.md)|
|[Save](Outlook.BusinessCardView.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.BusinessCardView.Application.md)|
|[CardSize](Outlook.BusinessCardView.CardSize.md)|
|[Class](Outlook.BusinessCardView.Class.md)|
|[Filter](Outlook.BusinessCardView.Filter.md)|
|[HeadingsFont](Outlook.BusinessCardView.HeadingsFont.md)|
|[Language](Outlook.BusinessCardView.Language.md)|
|[LockUserChanges](Outlook.BusinessCardView.LockUserChanges.md)|
|[Name](Outlook.BusinessCardView.Name.md)|
|[Parent](Outlook.BusinessCardView.Parent.md)|
|[SaveOption](Outlook.BusinessCardView.SaveOption.md)|
|[Session](Outlook.BusinessCardView.Session.md)|
|[SortFields](Outlook.BusinessCardView.SortFields.md)|
|[Standard](Outlook.BusinessCardView.Standard.md)|
|[ViewType](Outlook.BusinessCardView.ViewType.md)|
|[XML](Outlook.BusinessCardView.XML.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]