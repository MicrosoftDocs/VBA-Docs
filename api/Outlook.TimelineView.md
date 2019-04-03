---
title: TimelineView object (Outlook)
keywords: vbaol11.chm3185
f1_keywords:
- vbaol11.chm3185
ms.prod: outlook
api_name:
- Outlook.TimelineView
ms.assetid: fb14c1a1-f542-fa1e-f30f-c5ee3d2f0206
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineView object (Outlook)

Represents a view that displays Outlook items in a timeline.


## Remarks

The  **TimelineView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to display Outlook items within a timeline.

Outlook provides several built-in  **TimelineView** objects, and you can also create custom **TimelineView** objects. Use the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **TimelineView** to a **[Folder](Outlook.Folder.md)** object. Use the **[Standard](Outlook.TimelineView.Standard.md)** property to determine if an existing **TimelineView** object is built-in or custom.

The  **TimelineView** object supports several different view modes, depending on the desired layout and time period in which to display Outlook items. Use the **[TimelineViewMode](Outlook.TimelineView.TimelineViewMode.md)** property to set the view mode, the **[StartField](Outlook.TimelineView.StartField.md)** property to specify the Outlook item property that contains the start date, and the **[EndField](Outlook.TimelineView.EndField.md)** property to specify the Outlook item property that contains the end date for Outlook items to be displayed.

You can configure the appearance of the  **TimelineView**, depending on the view mode. Use the **[ShowWeekNumbers](Outlook.TimelineView.ShowWeekNumbers.md)** property to indicate whether week numbers are displayed in the time scale for the view. Use the **[UpperScaleFont](Outlook.TimelineView.UpperScaleFont.md)** and **[LowerScaleFont](Outlook.TimelineView.LowerScaleFont.md)** properties to specify the font used when displaying, respectively, the upper and lower portions of the time scale for the view.

You can also configure how Outlook items appear within the  **TimelineView** object. Use the **[ItemFont](Outlook.TimelineView.ItemFont.md)** property to specify the font used to display Outlook item labels and the **[MaxLabelWidth](Outlook.TimelineView.MaxLabelWidth.md)** property to specify the length of labels for Outlook items in the view. Use the **[DefaultExpandCollapseSetting](Outlook.TimelineView.DefaultExpandCollapseSetting.md)** property to determine if Outlook items are expanded by default in the view. Use the **[Filter](Outlook.TimelineView.Filter.md)** property to determine which Outlook items to display in the view and the **[GroupByFields](Outlook.TimelineView.GroupByFields.md)** collection to specify the Outlook item properties by which Outlook items are grouped in the view. If you set the **TimelineViewMode** to **olTimelineViewMonth**, you can use the **[ShowLabelWhenViewingByMonth](Outlook.TimelineView.ShowLabelWhenViewingByMonth.md)** property to determine if labels for Outlook items are displayed in the view.

The definition for each  **TimelineView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](Outlook.TimelineView.XML.md)** property to work with the XML definition for the **TimelineView** object.

Use the  **[Apply](Outlook.TimelineView.Apply.md)** method to apply any changes made to the **TimelineView** object to the current view. Use the **[Save](Outlook.TimelineView.Save.md)** method to persist any changes made to the **TimelineView** object. Use the **[LockUserChanges](Outlook.TimelineView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **TimelineView** objects, but you cannot delete them. Use the **[Delete](Outlook.TimelineView.Delete.md)** method to delete a custom **TimelineView** object. Use the **[Reset](Outlook.TimelineView.Reset.md)** method to reset the properties of a built-in **TimelineView** object to their default values.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **TimelineView** object to display Outlook items by month, with week number labels on the lower portion of the timeline scale, with labels no longer than 40 characters.


```vb
Private Sub ConfigureMonthTimelineView() 
 
 Dim objTimelineView As TimelineView 
 
 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olTimelineView Then 
 
 
 
 ' Obtain a TimelineView object reference for the 
 
 ' current timeline view. 
 
 Set objTimelineView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Configure the TimelineView object so that it displays 
 
 ' Outlook items by month and week, displaying labels 
 
 ' no larger than 40 characters for Outlook items 
 
 ' displayed in the view. 
 
 With objTimelineView 
 
 ' Display items by month. 
 
 .TimelineViewMode = olTimelineViewMonth 
 
 
 
 ' Display week numbers. If this value is 
 
 ' set to False when TimelineViewMode is 
 
 ' set to olTimelineViewMonth, the day 
 
 ' numbers are displayed instead. 
 
 .ShowWeekNumbers = True 
 
 
 
 ' Display labels for Outlook items 
 
 ' while TimelineViewMode is set to 
 
 ' olTimelineViewMonth. 
 
 .ShowLabelWhenViewingByMonth = True 
 
 
 
 ' Show no more than the first 40 characters 
 
 ' for each Outlook item in the view. 
 
 .MaxLabelWidth = 40 
 
 
 
 ' Save and apply the view. 
 
 .Save 
 
 .Apply 
 
 End With 
 
 End If 
 
 
 
End Sub 
 

```


## Methods



|Name|
|:-----|
|[Apply](Outlook.TimelineView.Apply.md)|
|[Copy](Outlook.TimelineView.Copy.md)|
|[Delete](Outlook.TimelineView.Delete.md)|
|[GoToDate](Outlook.TimelineView.GoToDate.md)|
|[Reset](Outlook.TimelineView.Reset.md)|
|[Save](Outlook.TimelineView.Save.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.TimelineView.Application.md)|
|[Class](Outlook.TimelineView.Class.md)|
|[DefaultExpandCollapseSetting](Outlook.TimelineView.DefaultExpandCollapseSetting.md)|
|[EndField](Outlook.TimelineView.EndField.md)|
|[Filter](Outlook.TimelineView.Filter.md)|
|[GroupByFields](Outlook.TimelineView.GroupByFields.md)|
|[ItemFont](Outlook.TimelineView.ItemFont.md)|
|[Language](Outlook.TimelineView.Language.md)|
|[LockUserChanges](Outlook.TimelineView.LockUserChanges.md)|
|[LowerScaleFont](Outlook.TimelineView.LowerScaleFont.md)|
|[MaxLabelWidth](Outlook.TimelineView.MaxLabelWidth.md)|
|[Name](Outlook.TimelineView.Name.md)|
|[Parent](Outlook.TimelineView.Parent.md)|
|[SaveOption](Outlook.TimelineView.SaveOption.md)|
|[Session](Outlook.TimelineView.Session.md)|
|[ShowLabelWhenViewingByMonth](Outlook.TimelineView.ShowLabelWhenViewingByMonth.md)|
|[ShowWeekNumbers](Outlook.TimelineView.ShowWeekNumbers.md)|
|[Standard](Outlook.TimelineView.Standard.md)|
|[StartField](Outlook.TimelineView.StartField.md)|
|[TimelineViewMode](Outlook.TimelineView.TimelineViewMode.md)|
|[UpperScaleFont](Outlook.TimelineView.UpperScaleFont.md)|
|[ViewType](Outlook.TimelineView.ViewType.md)|
|[XML](Outlook.TimelineView.XML.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]