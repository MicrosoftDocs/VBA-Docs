---
title: Categorize Your Outlook Items
ms.prod: outlook
ms.assetid: e8cfb450-b8b0-bee6-fdf0-d0a92bf9af56
ms.date: 06/08/2017
localization_priority: Normal
---


# Categorize Your Outlook Items

Microsoft Outlook provides color categorization functionality, in which Outlook items can be categorized and displayed by category. Multiple color categories can be applied to a single Outlook item, and Outlook items can be grouped or sorted by color category. Shortcut keys can be assigned to each color category, to allow users to more easily categorize items. Color categories are user-defined, and can be created, deleted, and changed either programmatically or by user action within the Outlook user interface.

The  **[Category](../../../api/Outlook.Category.md)** object represents a single user-defined color category in the Master Category List, the list of color categories presented in the Outlook user interface and represented by the **[Categories](../../../api/Outlook.NameSpace.Categories.md)** collection of the **[NameSpace](../../../api/Outlook.NameSpace.md)** object. **Category** objects are identified with a globally unique identifier (GUID) when created, and this identifier cannot be changed. However, the name, color, and shortcut key associated with a color category can be changed by setting the **[Name](../../../api/Outlook.Category.Name.md)**,  **[Color](../../../api/Outlook.Category.Color.md)**, and  **[ShortcutKey](../../../api/Outlook.Category.ShortcutKey.md)** properties, respectively, of the **Category** object. The **[CategoryID](../../../api/Outlook.Category.CategoryID.md)** property can be used to retrieve the identifier of a **Category** object.

## Assigning Categories to Outlook Items

Categories can be assigned to Outlook items by specifying the names of the appropriate  **Category** objects in a comma-delimited string in the **Categories** property of the following objects:



| **[AppointmentItem](../../../api/Outlook.AppointmentItem.md)**| **[RemoteItem](../../../api/Outlook.RemoteItem.md)**|
|:-----|:-----|
| **[ContactItem](../../../api/Outlook.ContactItem.md)**| **[ReportItem](../../../api/Outlook.ReportItem.md)**|
| **[DistListItem](../../../api/Outlook.DistListItem.md)**| **[SharingItem](../../../api/Outlook.SharingItem.md)**|
| **[DocumentItem](../../../api/Outlook.DocumentItem.md)**| **[PostItem](../../../api/Outlook.PostItem.md)**|
| **[JournalItem](../../../api/Outlook.JournalItem.md)**| **[TaskItem](../../../api/Outlook.TaskItem.md)**|
| **[MailItem](../../../api/Outlook.MailItem.md)**| **[TaskRequestAcceptItem](../../../api/Outlook.TaskRequestAcceptItem.md)**|
| **[MeetingItem](../../../api/Outlook.MeetingItem.md)**| **[TaskRequestDeclineItem](../../../api/Outlook.TaskRequestDeclineItem.md)**|
| **[MobileItem](../../../api/overview/Outlook.md)**| **[TaskRequestItem](../../../api/Outlook.TaskRequestItem.md)**|
| **[NoteItem](../../../api/Outlook.NoteItem.md)**| **[TaskRequestUpdateItem](../../../api/Outlook.TaskRequestUpdateItem.md)**|

Outlook items are displayed based on the category name stored in the  **Categories** property of that Outlook item. Because category names are stored as part of the Outlook item, it is possible to have a category name in an Outlook item that is not present in the Master Category List. For example, a category may have been removed.

If a  **Category** object with a corresponding **Name** property value does not exist in the **Categories** collection of the **NameSpace** object that contains the Outlook item, the category name associated with that Outlook item is still displayed, but without an associated color.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]