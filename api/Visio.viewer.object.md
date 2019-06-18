---
title: Viewer object (Visio Viewer)
ms.prod: visio
ms.assetid: 4d25251a-5c4d-42d4-a73e-7e1e987ff593
ms.date: 06/08/2017
localization_priority: Normal
---


# Viewer object (Visio Viewer)

The  **Viewer** object is a programmable ActiveX control that enables you to display Visio drawings (with limited functionality) on webpages and in Windows Forms, so that users who do not have Visio installed on their computers can view and interact with them.


## Remarks

With Visio Viewer, users can open, view, or print Visio drawings, even if they do not have Microsoft Visio 2013 installed. They cannot, however, edit or save drawings, or create a new Visio drawing. For that, they need to install Visio.

The  **Viewer** object is the entry point to the **Viewer** object model, and represents an instance of the Viewer control. The properties, events, and methods available in the **Viewer** object model let you load and unload Visio drawings in Visio Viewer, temporarily change properties and settings of the drawing, react to user input, and customize the Visio Viewer environment. In many cases, these members correspond to the options available to users in the Visio Viewer user interface (UI).

The following is a partial listing of the members of the  **Viewer** object and their functions and provides a sampling of the programming options available to developers. See the table of contents of this reference for the complete list of members. See [About Programming Visio Viewer](Visio.about.programming.md) for code samples that show how to get an instance of the **Viewer** object in the available development environments.

Use the  **[Load](Visio.Load.md)** method to load a Visio drawing into Visio Viewer, and use the **[Unload](Visio.Unload.md)** method to unload the drawing. You can also use the **[SRC](Visio.viewer.src.property.md)** property to get and set the file name and path for the current drawing.

Use the  **[DisplayAbout](Visio.DisplayAbout.md)**,  **[DisplayContextMenu](Visio.DisplayContextMenu.md)**,  **[DisplayHelp](Visio.DisplayHelp.md)**, and  **[DisplayPropertyDialog](Visio.DisplayPropertyDialog.md)** methods to display the dialog boxes and shortcut menus available in the Visio Viewer UI.

Use the  **[SelectShape](Visio.SelectShape.md)** method to select a particular shape in the drawing and the **[ShapeName](Visio.ShapeName.md)** and **[ShapeCount](Visio.ShapeCount.md)** properties to get information about shapes in the drawing.

Use properties such as  **[BackColor](Visio.viewer.backcolor.property.md)**,  **[GridVisible](Visio.GridVisible.md)**,  **[LayerColor](Visio.LayerColor.md)**,  **[PageColor](Visio.PageColor.md)**,  **[ScrollbarsVisible](Visio.ScrollbarsVisible.md)**, and  **[ToolbarVisible](Visio.ToolbarVisible.md)** to customize the appearance of the Visio Viewer UI.

Use the  **[CustomPropertyCount](Visio.CustomPropertyCount.md)**,  **[CustomPropertyName](Visio.CustomPropertyName.md)**, and  **[CustomPropertyValue](Visio.CustomPropertyValue.md)** properties to determine shape data (custom properties).

Use events such as  **[OnLayerChanged](Visio.OnLayerChanged.md)** and **[OnSelectionChanged](Visio.OnSelectionChanged.md)** to respond to user input.

## Events

- [OnDocumentLoaded](Visio.OnDocumentLoaded.md)
- [OnDocumentUnloaded](Visio.OnDocumentUnloaded.md)
- [OnLayerChanged](Visio.OnLayerChanged.md)
- [OnMarkupOverlaysVisibleChanged](Visio.OnMarkupOverlaysVisibleChanged.md)
- [OnPageChanged](Visio.OnPageChanged.md)
- [OnReviewerChanged](Visio.OnReviewerChanged.md)
- [OnSelectionChanged](Visio.OnSelectionChanged.md)
- [OnToolbarCustomized](Visio.OnToolbarCustomized.md)
- [OnViewChanged](Visio.OnViewChanged.md)

## Methods

- [DisplayAbout](Visio.DisplayAbout.md)
- [DisplayContextMenu](Visio.DisplayContextMenu.md)
- [DisplayHelp](Visio.DisplayHelp.md)
- [DisplayPropertyDialog](Visio.DisplayPropertyDialog.md)
- [FollowHyperlink](Visio.FollowHyperlink.md)
- [GetErrorMessage](Visio.GetErrorMessage.md)
- [GetPageView](Visio.GetPageView.md)
- [Load](Visio.Load.md)
- [Pan](Visio.Pan.md)
- [SelectShape](Visio.SelectShape.md)
- [SetPageView](Visio.SetPageView.md)
- [Unload](Visio.Unload.md)
- [ZoomToPoint](Visio.ZoomToPoint.md)
- [ZoomToRect](Visio.ZoomToRect.md)

## Properties

- [AlertsEnabled](Visio.AlertsEnabled.md)
- [BackColor](Visio.viewer.backcolor.property.md)
- [BuildNumber](Visio.BuildNumber.md)
- [ContextMenuEnabled](Visio.ContextMenuEnabled.md)
- [CurrentPageIndex](Visio.CurrentPageIndex.md)
- [CustomPropertyCount](Visio.CustomPropertyCount.md)
- [CustomPropertyName](Visio.CustomPropertyName.md)
- [CustomPropertyValue](Visio.CustomPropertyValue.md)
- [DocumentLoaded](Visio.DocumentLoaded.md)
- [GridVisible](Visio.GridVisible.md)
- [HighQualityRender](Visio.HighQualityRender.md)
- [HyperlinkAddress](Visio.HyperlinkAddress.md)
- [HyperlinkCount](Visio.HyperlinkCount.md)
- [LastErrorCode](Visio.LastErrorCode.md)
- [LayerColor](Visio.LayerColor.md)
- [LayerColorOverride](Visio.LayerColorOverride.md)
- [LayerColorTrans](Visio.LayerColorTrans.md)
- [LayerCount](Visio.LayerCount.md)
- [LayerDeleted](Visio.LayerDeleted.md)
- [LayerName](Visio.LayerName.md)
- [LayerVisible](Visio.LayerVisible.md)
- [MajorVersionNumber](Visio.MajorVersionNumber.md)
- [MarkupOverlaysVisible](Visio.MarkupOverlaysVisible.md)
- [MinorVersionNumber](Visio.MinorVersionNumber.md)
- [PageColor](Visio.PageColor.md)
- [PageCount](Visio.PageCount.md)
- [PageIDToIndex](Visio.PageIDToIndex.md)
- [PageIndexToID](Visio.PageIndexToID.md)
- [PageName](Visio.PageName.md)
- [PageTabsVisible](Visio.PageTabsVisible.md)
- [PageVisible](Visio.PageVisible.md)
- [ParentShape](Visio.ParentShape.md)
- [PropertyDialogEnabled](Visio.PropertyDialogEnabled.md)
- [ReviewerColor](Visio.ReviewerColor.md)
- [ReviewerCount](Visio.ReviewerCount.md)
- [ReviewerID](Visio.ReviewerID.md)
- [ReviewerInitial](Visio.ReviewerInitial.md)
- [ReviewerMarkupVisible](Visio.ReviewerMarkupVisible.md)
- [ReviewerName](Visio.ReviewerName.md)
- [ScrollbarsVisible](Visio.ScrollbarsVisible.md)
- [SelectedShapeIndex](Visio.SelectedShapeIndex.md)
- [ShapeAtPoint](Visio.ShapeAtPoint.md)
- [ShapeCount](Visio.ShapeCount.md)
- [ShapeIDToIndex](Visio.ShapeIDToIndex.md)
- [ShapeIndexToID](Visio.ShapeIndexToID.md)
- [ShapeName](Visio.ShapeName.md)
- [SRC](Visio.viewer.src.property.md)
- [SubShapeAtPoint](Visio.SubShapeAtPoint.md)
- [ToolbarButtons](Visio.ToolbarButtons.md)
- [ToolbarCustomizable](Visio.ToolbarCustomizable.md)
- [ToolbarVisible](Visio.ToolbarVisible.md)
- [Zoom](Visio.viewer.zoom.property.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]