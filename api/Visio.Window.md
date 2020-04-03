---
title: Window object (Visio)
keywords: vis_sdr.chm10305
f1_keywords:
- vis_sdr.chm10305
ms.prod: visio
api_name:
- Visio.Window
ms.assetid: 5b49eb0f-07ea-00c7-52f1-2a3115a4b8ae
ms.date: 06/19/2019
localization_priority: Normal
---


# Window object (Visio)

Represents an open window in a Microsoft Visio instance.


## Remarks

The default property of a **Window** object is **Application**.

To retrieve the active window in an instance of Visio, use the **[ActiveWindow](visio.application.activewindow.md)** property of an **Application** object.
    
To retrieve a **Page** object that represents the page shown in the window, use the **Page** property of a **Window** object.
    
To retrieve a **Document** object that represents the document displayed in that window, use the **Document** property.
    
To retrieve a **Selection** object that represents the shapes selected in that window, use the **Selection** property.
    
> [!NOTE] 
> Beginning with Microsoft Visio 2002, the following methods of the **Window** object are obsolete: **AddToGroup**, **Cut**, **Combine**, **Copy**, **Delete**, **Duplicate**, **Fragment**, **Group**, **Intersect**, **Join**, **RemoveFromGroup**, **Subtract**, **Trim**, and **Union**. Existing solutions that invoke these methods will continue to work properly; however, new or rebuilt solutions should use these methods with the **Selection** object.

In addition, the **Window** object's **Paste** method is now obsolete. Use the **Paste** or **PasteSpecial** method of the **Page**, **Master**, or **Shape** object (use the **Shape** object in the case of group shapes).


## Events

- [BeforeWindowClosed](Visio.Window.BeforeWindowClosed.md)
- [BeforeWindowPageTurn](Visio.Window.BeforeWindowPageTurn.md)
- [BeforeWindowSelDelete](Visio.Window.BeforeWindowSelDelete.md)
- [KeyDown](Visio.Window.KeyDown.md)
- [KeyPress](Visio.Window.KeyPress.md)
- [KeyUp](Visio.Window.KeyUp.md)
- [MouseDown](Visio.Window.MouseDown.md)
- [MouseMove](Visio.Window.MouseMove.md)
- [MouseUp](Visio.Window.MouseUp.md)
- [OnKeystrokeMessageForAddon](Visio.Window.OnKeystrokeMessageForAddon.md)
- [QueryCancelWindowClose](Visio.Window.QueryCancelWindowClose.md)
- [SelectionChanged](Visio.Window.SelectionChanged.md)
- [ViewChanged](Visio.Window.ViewChanged.md)
- [WindowActivated](Visio.Window.WindowActivated.md)
- [WindowChanged](Visio.Window.WindowChanged.md)
- [WindowCloseCanceled](Visio.Window.WindowCloseCanceled.md)
- [WindowTurnedToPage](Visio.Window.WindowTurnedToPage.md)

## Methods

- [Activate](Visio.Window.Activate.md)
- [CenterViewOnShape](Visio.Window.CenterViewOnShape.md)
- [Close](Visio.Window.Close.md)
- [DeselectAll](Visio.Window.DeselectAll.md)
- [DockedStencils](Visio.Window.DockedStencils.md)
- [GetViewRect](Visio.Window.GetViewRect.md)
- [GetWindowRect](Visio.Window.GetWindowRect.md)
- [Group](Visio.Window.Group.md)
- [NewWindow](Visio.Window.NewWindow.md)
- [Paste](Visio.Window.Paste.md)
- [Scroll](Visio.Window.Scroll.md)
- [ScrollViewTo](Visio.Window.ScrollViewTo.md)
- [Select](Visio.Window.Select.md)
- [SelectAll](Visio.Window.SelectAll.md)
- [SetViewRect](Visio.Window.SetViewRect.md)
- [SetWindowRect](Visio.Window.SetWindowRect.md)

## Properties

- [AllowEditing](Visio.Window.AllowEditing.md)
- [Application](Visio.Window.Application.md)
- [BackgroundColor](Visio.Window.BackgroundColor.md)
- [BackgroundColorGradient](Visio.Window.BackgroundColorGradient.md)
- [Caption](Visio.Window.Caption.md)
- [Document](Visio.Window.Document.md)
- [EventList](Visio.Window.EventList.md)
- [ID](Visio.Window.ID.md)
- [Index](Visio.Window.Index.md)
- [InPlace](Visio.Window.InPlace.md)
- [IsEditingOLE](Visio.Window.IsEditingOLE.md)
- [IsEditingText](Visio.Window.IsEditingText.md)
- [Master](Visio.Window.Master.md)
- [MasterShortcut](Visio.Window.MasterShortcut.md)
- [MergeCaption](Visio.Window.MergeCaption.md)
- [MergeClass](Visio.Window.MergeClass.md)
- [MergeID](Visio.Window.MergeID.md)
- [MergePosition](Visio.Window.MergePosition.md)
- [ObjectType](Visio.Window.ObjectType.md)
- [Page](Visio.Window.Page.md)
- [PageTabWidth](Visio.Window.PageTabWidth.md)
- [Parent](Visio.Window.Parent.md)
- [ParentWindow](Visio.Window.ParentWindow.md)
- [PersistsEvents](Visio.Window.PersistsEvents.md)
- [ReviewerMarkupVisible](Visio.Window.ReviewerMarkupVisible.md)
- [ScrollLock](Visio.Window.ScrollLock.md)
- [SelectedCell](Visio.Window.SelectedCell.md)
- [SelectedDataRecordset](Visio.Window.SelectedDataRecordset.md)
- [SelectedDataRowID](Visio.Window.SelectedDataRowID.md)
- [SelectedMasters](Visio.Window.SelectedMasters.md)
- [SelectedText](Visio.Window.SelectedText.md)
- [SelectedValidationIssue](Visio.Window.SelectedValidationIssue.md)
- [Selection](Visio.Window.Selection.md)
- [SelectionForDragCopy](Visio.Window.SelectionForDragCopy.md)
- [Shape](Visio.Window.Shape.md)
- [ShowConnectPoints](Visio.Window.ShowConnectPoints.md)
- [ShowGrid](Visio.Window.ShowGrid.md)
- [ShowGuides](Visio.Window.ShowGuides.md)
- [ShowPageBreaks](Visio.Window.ShowPageBreaks.md)
- [ShowPageOutline](Visio.Window.ShowPageOutline.md)
- [ShowPageTabs](Visio.Window.ShowPageTabs.md)
- [ShowRulers](Visio.Window.ShowRulers.md)
- [ShowScrollBars](Visio.Window.ShowScrollBars.md)
- [Stat](Visio.Window.Stat.md)
- [SubType](Visio.Window.SubType.md)
- [Type](Visio.Window.Type.md)
- [ViewFit](Visio.Window.ViewFit.md)
- [Visible](Visio.Window.Visible.md)
- [WindowHandle32](Visio.Window.WindowHandle32.md)
- [Windows](Visio.Window.Windows.md)
- [WindowState](Visio.Window.WindowState.md)
- [Zoom](Visio.Window.Zoom.md)
- [ZoomBehavior](Visio.Window.ZoomBehavior.md)
- [ZoomLock](Visio.Window.ZoomLock.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]