---
title: TableView object (Outlook)
keywords: vbaol11.chm3204
f1_keywords:
- vbaol11.chm3204
ms.prod: outlook
api_name:
- Outlook.TableView
ms.assetid: 026e27f8-1655-060d-e8cc-87eaaf4f1510
ms.date: 06/08/2017
localization_priority: Normal
---


# TableView object (Outlook)

Represents a view that displays Outlook items in a table, with each item in a row and the details of the item in the columns.


## Remarks

The **TableView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to display Outlook items in a table.

Outlook provides several built-in  **TableView** objects, and you can also create custom **TableView** objects. Use the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **TableView** to a **[Folder](Outlook.Folder.md)** object. Use the **Standard** property to determine if an existing **TableView** object is built-in or custom.

You can configure the appearance and functionality of the  **TableView** object. Use the **[AutomaticColumnSizing](Outlook.TableView.AutomaticColumnSizing.md)** property to determine whether the view automatically resizes columns and the **[AutomaticGrouping](Outlook.TableView.AutomaticGrouping.md)** property to determine if the view automatically groups Outlook items. Use the **[AutoPreview](Outlook.TableView.AutoPreview.md)** property to determine whether preview information is displayed within the row for an Outlook item in the view, and the **[AutoPreviewFont](Outlook.TableView.AutoPreviewFont.md)** property to specify the font used to display preview information. Use the **[Multiline](Outlook.TableView.Multiline.md)** property to determine whether to show Outlook items in multiline mode.

You can also configure how Outlook items appear within the  **TableView** object. Use the **[ColumnFont](Outlook.TableView.ColumnFont.md)** property to specify the font used for column headers and the **[RowFont](Outlook.TableView.RowFont.md)** property to specify the font used for Outlook items in the view. Use the **[AllowInCellEditing](Outlook.TableView.AllowInCellEditing.md)** property to allow editing of Outlook item property values in the view. Use the **[Filter](Outlook.TableView.Filter.md)** property to determine which Outlook items to display in the view and the **[ViewFields](Outlook.TableView.ViewFields.md)** collection to specify the Outlook item properties to display for each Outlook item. Use the **[GroupByFields](Outlook.TableView.GroupByFields.md)** to specify the Outlook item properties by which Outlook items are grouped, and the **[SortFields](Outlook.TableView.SortFields.md)** collection to specify the Outlook item properties by which Outlook items are sorted in the view.

The definition for each  **TableView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](Outlook.TableView.XML.md)** property to work with the XML definition for the **TableView** object.

Use the  **[Apply](Outlook.TableView.Apply.md)** method to apply any changes made to the **TableView** object to the current view. Use the **[Save](Outlook.TableView.Save.md)** method to persist any changes made to the **TableView** object. Use the **[LockUserChanges](Outlook.TableView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **TableView** objects, but you cannot delete them. Use the **[Delete](Outlook.TableView.Delete.md)** method to delete a custom **TableView** object. Use the **[Reset](Outlook.TableView.Reset.md)** method to reset the properties of a built-in **TableView** object to their default values.


## Methods



|Name|
|:-----|
|[Apply](Outlook.TableView.Apply.md)|
|[Copy](Outlook.TableView.Copy.md)|
|[Delete](Outlook.TableView.Delete.md)|
|[GetTable](Outlook.TableView.GetTable.md)|
|[GoToDate](Outlook.TableView.GoToDate.md)|
|[Reset](Outlook.TableView.Reset.md)|
|[Save](Outlook.TableView.Save.md)|

## Properties



|Name|
|:-----|
|[AllowInCellEditing](Outlook.TableView.AllowInCellEditing.md)|
|[AlwaysExpandConversation](Outlook.TableView.AlwaysExpandConversation.md)|
|[Application](Outlook.TableView.Application.md)|
|[AutoFormatRules](Outlook.TableView.AutoFormatRules.md)|
|[AutomaticColumnSizing](Outlook.TableView.AutomaticColumnSizing.md)|
|[AutomaticGrouping](Outlook.TableView.AutomaticGrouping.md)|
|[AutoPreview](Outlook.TableView.AutoPreview.md)|
|[AutoPreviewFont](Outlook.TableView.AutoPreviewFont.md)|
|[Class](Outlook.TableView.Class.md)|
|[ColumnFont](Outlook.TableView.ColumnFont.md)|
|[DefaultExpandCollapseSetting](Outlook.TableView.DefaultExpandCollapseSetting.md)|
|[Filter](Outlook.TableView.Filter.md)|
|[GridLineStyle](Outlook.TableView.GridLineStyle.md)|
|[GroupByFields](Outlook.TableView.GroupByFields.md)|
|[HideReadingPaneHeaderInfo](Outlook.TableView.HideReadingPaneHeaderInfo.md)|
|[Language](Outlook.TableView.Language.md)|
|[LockUserChanges](Outlook.TableView.LockUserChanges.md)|
|[MaxLinesInMultiLineView](Outlook.TableView.MaxLinesInMultiLineView.md)|
|[MultiLine](Outlook.TableView.Multiline.md)|
|[MultiLineWidth](Outlook.TableView.MultiLineWidth.md)|
|[Name](Outlook.TableView.Name.md)|
|[Parent](Outlook.TableView.Parent.md)|
|[RowFont](Outlook.TableView.RowFont.md)|
|[SaveOption](Outlook.TableView.SaveOption.md)|
|[Session](Outlook.TableView.Session.md)|
|[ShowConversationByDate](Outlook.TableView.ShowConversationByDate.md)|
|[ShowConversationSendersAboveSubject](Outlook.TableView.ShowConversationSendersAboveSubject.md)|
|[ShowFullConversations](Outlook.TableView.ShowFullConversations.md)|
|[ShowItemsInGroups](Outlook.TableView.ShowItemsInGroups.md)|
|[ShowNewItemRow](Outlook.TableView.ShowNewItemRow.md)|
|[ShowReadingPane](Outlook.TableView.ShowReadingPane.md)|
|[SortFields](Outlook.TableView.SortFields.md)|
|[Standard](Outlook.TableView.Standard.md)|
|[ViewFields](Outlook.TableView.ViewFields.md)|
|[ViewType](Outlook.TableView.ViewType.md)|
|[XML](Outlook.TableView.XML.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]