---
title: Outlook Views
ms.prod: outlook
ms.assetid: cbaa3192-6c27-26c0-ebd6-f6489c2e812e
ms.date: 06/08/2019
localization_priority: Normal
---


# Outlook Views

Microsoft Outlook allows you to create customizable views that allow you to better sort, group, and ultimately view data of all different types within an explorer. There are a variety of different view types that provide the flexibility needed to create and maintain your important data. Outlook supports the following objects, derived from the **[View](../../../api/Outlook.View.md)** object.


|**Object name**|**Description**|
|:-----|:-----|
| **[BusinessCardView](../../../api/Outlook.CardView.md)**|This object allows you to view data as a series of Electronic Business Card (EBC) images.|
| **[CalendarView](../../../api/Outlook.CalendarView.md)**|This object allows you to view data in a calendar format.|
| **[CardView](../../../api/Outlook.CardView.md)**|This object allows you to view data in a series of cards.|
| **[IconView](../../../api/Outlook.IconView.md)**|This object allows you to view data as icons, similar to a Windows folder or explorer.|
| **[TableView](../../../api/Outlook.TableView.md)**|This object allows you to view data in a simple, field-based table.|
| **[TimelineView](../../../api/Outlook.TimelineView.md)**|This object allows you to view data in a customizable linear time line.|

While you can use the **View** object to interact with the properties and methods common to all views, you must cast the **View** object to one of the derived view objects, such as the **CardView** object, to access certain properties, such as the **[HeadingsFont](../../../api/Outlook.CardView.HeadingsFont.md)** property of the **CardView** object. You can use the **[ViewType](../../../api/Outlook.View.ViewType.md)** property of the **View** object to determine which type of view is represented by that object.

You can define a new view by using the **[Add](../../../api/Outlook.Views.Add.md)** method of the **[Views](../../../api/Outlook.Views.md)** collection for a **[Folder](../../../api/Outlook.Folder.md)** object. Visibility for the view can be set either at the time of creation, by specifying an **[OlViewSaveOption](../../../api/Outlook.OlViewSaveOption.md)** constant in the _SaveOption_ parameter of the **Add** method, or any time after the view is created, by specifying an **OlViewSaveOption** constant for the **[SaveOption](../../../api/Outlook.View.SaveOption.md)** property of the **View** object. Adding a new view raises the **[ViewAdd](../../../api/Outlook.Views.ViewAdd.md)** event of the **Views** collection.
You can use the **[Remove](../../../api/Outlook.Views.Remove.md)** method of the **Views** object to remove an existing custom view. Removing a view raises the **[ViewRemove](../../../api/Outlook.ViewRemove.md)** event of the **Views** collection.
Once a view is defined, you can customize the view programmatically by casting the **View** object to one of the derived view objects, such as the **BusinessCardView** object, and performing whatever changes are needed. Use the **[Save](../../../api/Outlook.View.Save.md)** method of the derived view object or the **View** object to save any changes to the view.
You can apply the view, once defined and customized, to the current **[Explorer](../../../api/Outlook.Explorer.md)** object by using the **[Apply](../../../api/Outlook.View.Apply.md)** method of the derived view object or the **View** object. Applying a view raises the **[ViewSwitch](../../../api/Outlook.Explorer.ViewSwitch.md)** event of the **Explorer** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]