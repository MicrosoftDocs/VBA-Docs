---
title: Viewer object (Visio Viewer)
ms.prod: visio
ms.assetid: 4d25251a-5c4d-42d4-a73e-7e1e987ff593
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer object (Visio Viewer)

The **Viewer** object is a programmable ActiveX control that enables you to display Visio drawings (with limited functionality) on webpages and in Windows Forms, so that users who do not have Visio installed on their computers can view and interact with them.


## Remarks

With Visio Viewer, users can open, view, or print Visio drawings, even if they do not have Microsoft Visio installed. They cannot, however, edit or save drawings, or create a new Visio drawing. For that, they need to install Visio.

The **Viewer** object is the entry point to the **Viewer** object model, and represents an instance of the Viewer control. The properties, events, and methods available in the **Viewer** object model let you load and unload Visio drawings in Visio Viewer, temporarily change properties and settings of the drawing, react to user input, and customize the Visio Viewer environment. In many cases, these members correspond to the options available to users in the Visio Viewer user interface (UI).

The following is a partial listing of the members of the **Viewer** object and their functions and provides a sampling of the programming options available to developers. For code samples that show how to get an instance of the **Viewer** object in the available development environments, see [About Programming Visio Viewer](Visio.ViewerRef.AboutProgramming.md). 

Use the **Load** method to load a Visio drawing into Visio Viewer, and use the **Unload** method to unload the drawing. You can also use the **Src** property to get and set the file name and path for the current drawing.

Use the **DisplayAbout**, **DisplayContextMenu**, **DisplayHelp**, and **DisplayPropertyDialog** methods to display the dialog boxes and shortcut menus available in the Visio Viewer UI.

Use the **SelectShape** method to select a particular shape in the drawing, and use the **ShapeName** and **ShapeCount** properties to get information about shapes in the drawing.

Use properties such as **BackColor**, **GridVisible**, **LayerColor**, **PageColor**, **ScrollbarsVisible**, and **ToolbarVisible** to customize the appearance of the Visio Viewer UI.

Use the **CustomPropertyCount**, **CustomPropertyName**, and **CustomPropertyValue** properties to determine shape data (custom properties).

Use events such as **OnLayerChanged** and **OnSelectionChanged** to respond to user input.

## Events

- [OnDocumentLoaded](Visio.Viewer.OnDocumentLoaded.md)
- [OnDocumentUnloaded](Visio.Viewer.OnDocumentUnloaded.md)
- [OnLayerChanged](Visio.Viewer.OnLayerChanged.md)
- [OnMarkupOverlaysVisibleChanged](Visio.Viewer.OnMarkupOverlaysVisibleChanged.md)
- [OnPageChanged](Visio.Viewer.OnPageChanged.md)
- [OnReviewerChanged](Visio.Viewer.OnReviewerChanged.md)
- [OnSelectionChanged](Visio.Viewer.OnSelectionChanged.md)
- [OnToolbarCustomized](Visio.Viewer.OnToolbarCustomized.md)
- [OnViewChanged](Visio.Viewer.OnViewChanged.md)


## Methods

- [DisplayAbout](Visio.Viewer.DisplayAbout.md)
- [DisplayContextMenu](Visio.Viewer.DisplayContextMenu.md)
- [DisplayHelp](Visio.Viewer.DisplayHelp.md)
- [DisplayPropertyDialog](Visio.Viewer.DisplayPropertyDialog.md)
- [FollowHyperlink](Visio.Viewer.FollowHyperlink.md)
- [GetErrorMessage](Visio.Viewer.GetErrorMessage.md)
- [GetPageView](Visio.Viewer.GetPageView.md)
- [Load](Visio.Viewer.Load.md)
- [Pan](Visio.Viewer.Pan.md)
- [SelectShape](Visio.Viewer.SelectShape.md)
- [SetPageView](Visio.Viewer.SetPageView.md)
- [Unload](Visio.Viewer.Unload.md)
- [ZoomToPoint](Visio.Viewer.ZoomToPoint.md)
- [ZoomToRect](Visio.Viewer.ZoomToRect.md)


## Properties

- [AlertsEnabled](Visio.Viewer.AlertsEnabled.md)
- [BackColor](Visio.Viewer.Backcolor.md)
- [BuildNumber](Visio.Viewer.BuildNumber.md)
- [ContextMenuEnabled](Visio.Viewer.ContextMenuEnabled.md)
- [CurrentPageIndex](Visio.Viewer.CurrentPageIndex.md)
- [CustomPropertyCount](Visio.Viewer.CustomPropertyCount.md)
- [CustomPropertyName](Visio.Viewer.CustomPropertyName.md)
- [CustomPropertyValue](Visio.Viewer.CustomPropertyValue.md)
- [DocumentLoaded](Visio.Viewer.DocumentLoaded.md)
- [GridVisible](Visio.Viewer.GridVisible.md)
- [HighQualityRender](Visio.Viewer.HighQualityRender.md)
- [HyperlinkAddress](Visio.Viewer.HyperlinkAddress.md)
- [HyperlinkCount](Visio.Viewer.HyperlinkCount.md)
- [LastErrorCode](Visio.Viewer.LastErrorCode.md)
- [LayerColor](Visio.Viewer.LayerColor.md)
- [LayerColorOverride](Visio.Viewer.LayerColorOverride.md)
- [LayerColorTrans](Visio.Viewer.LayerColorTrans.md)
- [LayerCount](Visio.Viewer.LayerCount.md)
- [LayerDeleted](Visio.Viewer.LayerDeleted.md)
- [LayerName](Visio.Viewer.LayerName.md)
- [LayerVisible](Visio.Viewer.LayerVisible.md)
- [MajorVersionNumber](Visio.Viewer.MajorVersionNumber.md)
- [MarkupOverlaysVisible](Visio.Viewer.MarkupOverlaysVisible.md)
- [MinorVersionNumber](Visio.Viewer.MinorVersionNumber.md)
- [PageColor](Visio.Viewer.PageColor.md)
- [PageCount](Visio.Viewer.PageCount.md)
- [PageIDToIndex](Visio.Viewer.PageIDToIndex.md)
- [PageIndexToID](Visio.Viewer.PageIndexToID.md)
- [PageName](Visio.Viewer.PageName.md)
- [PageTabsVisible](Visio.Viewer.PageTabsVisible.md)
- [PageVisible](Visio.Viewer.PageVisible.md)
- [ParentShape](Visio.Viewer.ParentShape.md)
- [PropertyDialogEnabled](Visio.Viewer.PropertyDialogEnabled.md)
- [ReviewerColor](Visio.Viewer.ReviewerColor.md)
- [ReviewerCount](Visio.Viewer.ReviewerCount.md)
- [ReviewerID](Visio.Viewer.ReviewerID.md)
- [ReviewerInitial](Visio.Viewer.ReviewerInitial.md)
- [ReviewerMarkupVisible](Visio.Viewer.ReviewerMarkupVisible.md)
- [ReviewerName](Visio.Viewer.ReviewerName.md)
- [ScrollbarsVisible](Visio.Viewer.ScrollbarsVisible.md)
- [SelectedShapeIndex](Visio.Viewer.SelectedShapeIndex.md)
- [ShapeAtPoint](Visio.Viewer.ShapeAtPoint.md)
- [ShapeCount](Visio.Viewer.ShapeCount.md)
- [ShapeIDToIndex](Visio.Viewer.ShapeIDToIndex.md)
- [ShapeIndexToID](Visio.Viewer.ShapeIndexToID.md)
- [ShapeName](Visio.Viewer.ShapeName.md)
- [Src](Visio.Viewer.Src.md)
- [SubShapeAtPoint](Visio.Viewer.SubShapeAtPoint.md)
- [ToolbarButtons](Visio.Viewer.ToolbarButtons.md)
- [ToolbarCustomizable](Visio.Viewer.ToolbarCustomizable.md)
- [ToolbarVisible](Visio.Viewer.ToolbarVisible.md)
- [Zoom](Visio.Viewer.Zoom.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]