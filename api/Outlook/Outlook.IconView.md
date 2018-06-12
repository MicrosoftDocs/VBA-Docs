---
title: IconView Object (Outlook)
keywords: vbaol11.chm3206
f1_keywords:
- vbaol11.chm3206
ms.prod: outlook
api_name:
- Outlook.IconView
ms.assetid: dc2efa6c-4752-f713-f77e-378036f358dc
ms.date: 06/08/2017
---


# IconView Object (Outlook)

Represents a view that displays Outlook items as a series of labeled icons.


## Remarks

The  **IconView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to display Outlook items as large or small icons, with labels.

Outlook provides several built-in views, and you can also create custom  **IconView** objects. Use the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **IconView** to a **[Folder](Outlook.Folder.md)** object. Use the **[Standard](Outlook.IconView.Standard.md)** property to determine if an existing **IconView** object is built-in or custom.

The  **IconView** object supports several different view types, depending on the desired layout in which to display Outlook items. Use the **[IconViewType](Outlook.IconView.IconViewType.md)** property to set the view type.

You can also configure how Outlook items appear within the  **IconView** object. Use the **[IconPlacement](Outlook.IconView.IconPlacement.md)** property to determine how the icons for Outlook items are arranged within the view. Use the **[Filter](Outlook.IconView.Filter.md)** property to determine which Outlook items to display in the view and the **[SortFields](Outlook.IconView.SortFields.md)** collection to specify the Outlook item properties by which Outlook items are sorted in the view.

The definition for each  **IconView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](Outlook.IconView.XML.md)** property to work with the XML definition for the **IconView** object.

Use the  **[Apply](Outlook.IconView.Apply.md)** method to apply any changes made to the **IconView** object to the current view. Use the **[Save](Outlook.IconView.Save.md)** method to persist any changes made to the **IconView** object. Use the **[LockUserChanges](Outlook.IconView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **IconView** objects, but you cannot delete them. Use the **[Delete](Outlook.IconView.Delete.md)** method to delete a custom **IconView** object. Use the **[Reset](Outlook.IconView.Reset.md)** method to reset the properties of a built-in **IconView** object to their default values.


## Methods



|**Name**|
|:-----|
|[Apply](Outlook.IconView.Apply.md)|
|[Copy](Outlook.IconView.Copy.md)|
|[Delete](Outlook.IconView.Delete.md)|
|[GoToDate](Outlook.IconView.GoToDate.md)|
|[Reset](Outlook.IconView.Reset.md)|
|[Save](Outlook.IconView.Save.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.IconView.Application.md)|
|[Class](Outlook.IconView.Class.md)|
|[Filter](Outlook.IconView.Filter.md)|
|[IconPlacement](Outlook.IconView.IconPlacement.md)|
|[IconViewType](Outlook.IconView.IconViewType.md)|
|[Language](iconview-language-property-outlook.md)|
|[LockUserChanges](Outlook.IconView.LockUserChanges.md)|
|[Name](iconview-name-property-outlook.md)|
|[Parent](iconview-parent-property-outlook.md)|
|[SaveOption](iconview-saveoption-property-outlook.md)|
|[Session](iconview-session-property-outlook.md)|
|[SortFields](Outlook.IconView.SortFields.md)|
|[Standard](Outlook.IconView.Standard.md)|
|[ViewType](iconview-viewtype-property-outlook.md)|
|[XML](Outlook.IconView.XML.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
