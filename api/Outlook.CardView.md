---
title: CardView object (Outlook)
keywords: vbaol11.chm3207
f1_keywords:
- vbaol11.chm3207
ms.prod: outlook
api_name:
- Outlook.CardView
ms.assetid: cdac229b-f2b6-9ecb-e1a7-b53509426570
ms.date: 06/08/2017
localization_priority: Normal
---


# CardView object (Outlook)

Represents a view that displays Outlook items as a series of index cards.


## Remarks

The  **CardView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to display Outlook items as index cards.

Outlook provides several built-in  **CardView** objects, and you can also create custom **CardView** objects. Use the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **CardView** to a **[Folder](Outlook.Folder.md)** object. Use the **[Standard](Outlook.CardView.Standard.md)** property to determine if an existing **CardView** object is built-in or custom.

You can configure how Outlook items appear within the  **CardView** object. Use the **[MultiLineFieldHeight](Outlook.CardView.MultiLineFieldHeight.md)** property to specify the number of lines used to display multi-line text in each card, the **[HeadingsFont](Outlook.CardView.HeadingsFont.md)** property to specify the font used to display heading text on each card, and the **[BodyFont](Outlook.CardView.BodyFont.md)** property to specify the font used to display body text on each card. Use the **[AllowInCellEditing](Outlook.CardView.AllowInCellEditing.md)** property to allow editing of Outlook item property values in the view, and the **[ShowEmptyFields](Outlook.CardView.ShowEmptyFields.md)** property to display empty Outlook item properties in the view. Use the **[Filter](Outlook.CardView.Filter.md)** property to determine which Outlook items to display in the view, the **[ViewFields](Outlook.CardView.ViewFields.md)** collection to specify the Outlook item properties to display in each card, and the **[SortFields](Outlook.CardView.SortFields.md)** collection to specify the Outlook item properties by which Outlook items are sorted in the view.

The definition for each  **CardView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](Outlook.CardView.XML.md)** property to work with the XML definition for the **CardView** object.

Use the  **[Apply](Outlook.CardView.Apply.md)** method to apply any changes made to the **CardView** object to the current view. Use the **[Save](Outlook.CardView.Save.md)** method to persist any changes made to the **CardView** object. Use the **[LockUserChanges](Outlook.CardView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **CardView** objects, but you cannot delete them. Use the **[Delete](Outlook.CardView.Delete.md)** method to delete a custom **CardView** object. Use the **[Reset](Outlook.CardView.Reset.md)** method to reset the properties of a built-in **CardView** object to their default values.


## Methods



|Name|
|:-----|
|[Apply](Outlook.CardView.Apply.md)|
|[Copy](Outlook.CardView.Copy.md)|
|[Delete](Outlook.CardView.Delete.md)|
|[GoToDate](Outlook.CardView.GoToDate.md)|
|[Reset](Outlook.CardView.Reset.md)|
|[Save](Outlook.CardView.Save.md)|

## Properties



|Name|
|:-----|
|[AllowInCellEditing](Outlook.CardView.AllowInCellEditing.md)|
|[Application](Outlook.CardView.Application.md)|
|[AutoFormatRules](Outlook.CardView.AutoFormatRules.md)|
|[BodyFont](Outlook.CardView.BodyFont.md)|
|[Class](Outlook.CardView.Class.md)|
|[Filter](Outlook.CardView.Filter.md)|
|[HeadingsFont](Outlook.CardView.HeadingsFont.md)|
|[Language](Outlook.CardView.Language.md)|
|[LockUserChanges](Outlook.CardView.LockUserChanges.md)|
|[MultiLineFieldHeight](Outlook.CardView.MultiLineFieldHeight.md)|
|[Name](Outlook.CardView.Name.md)|
|[Parent](Outlook.CardView.Parent.md)|
|[SaveOption](Outlook.CardView.SaveOption.md)|
|[Session](Outlook.CardView.Session.md)|
|[ShowEmptyFields](Outlook.CardView.ShowEmptyFields.md)|
|[SortFields](Outlook.CardView.SortFields.md)|
|[Standard](Outlook.CardView.Standard.md)|
|[ViewFields](Outlook.CardView.ViewFields.md)|
|[ViewType](Outlook.CardView.ViewType.md)|
|[Width](Outlook.CardView.Width.md)|
|[XML](Outlook.CardView.XML.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]