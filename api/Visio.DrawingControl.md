---
title: DrawingControl object (Visio)
keywords: vis_sdr.chm0
f1_keywords:
- vis_sdr.chm0
ms.prod: visio
api_name:
- Visio.DrawingControl
ms.assetid: ad7c6abf-5bbd-5b84-4a63-eceaf90991a8
ms.date: 06/19/2019
localization_priority: Normal
---


# DrawingControl object (Visio)

A programmable ActiveX control that enables you to build Microsoft Visio functionality into programs that you create in Microsoft Visual Studio and other development platforms.

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

Use the **Document** property to get the **[Document](visio.document.md)** object associated with the instance of the Microsoft Visio Drawing Control and thereby gain access to the Visio object model.

Use the **HostID** property to assign a GUID or other string representation of the container application to a registry key.

Use the **NegotiateMenus** and **Negotiate Toolbars** properties to determine whether Visio menus and toolbars are merged with those of the host container application in the Visio Drawing Control, and to enable programmatic customizing of Visio menus and toolbars.

Use the **PageSizingBehavior** property to specify how the behavior of the control changes as the control is resized, with respect to the drawing page and any shapes on it.

Use the **Src** property to specify the Visio drawing to appear in the Visio Drawing Control.

Use the **Window** property to get the **[Window](visio.window.md)** object associated with the instance of the Visio Drawing Control and thereby gain access to the Visio object model.

The **DrawingControl** object has no default property.

## Events

- [AfterRemoveHiddenInformation](Visio.DrawingControl.AfterRemoveHiddenInformation.md)
- [BeforeDataRecordsetDelete](Visio.DrawingControl.BeforeDataRecordsetDelete.md)
- [BeforeDocumentClose](Visio.DrawingControl.BeforeDocumentClose.md)
- [BeforeDocumentSave](Visio.DrawingControl.BeforeDocumentSave.md)
- [BeforeDocumentSaveAs](Visio.DrawingControl.BeforeDocumentSaveAs.md)
- [BeforeMasterDelete](Visio.DrawingControl.BeforeMasterDelete.md)
- [BeforePageDelete](Visio.DrawingControl.BeforePageDelete.md)
- [BeforeSelectionDelete](Visio.DrawingControl.BeforeSelectionDelete.md)
- [BeforeShapeTextEdit](Visio.DrawingControl.BeforeShapeTextEdit.md)
- [BeforeStyleDelete](Visio.DrawingControl.BeforeStyleDelete.md)
- [BeforeWindowClosed](Visio.DrawingControl.BeforeWindowClosed.md)
- [BeforeWindowPageTurn](Visio.DrawingControl.BeforeWindowPageTurn.md)
- [BeforeWindowSelDelete](Visio.DrawingControl.BeforeWindowSelDelete.md)
- [ConvertToGroupCanceled](Visio.DrawingControl.ConvertToGroupCanceled.md)
- [DataRecordsetAdded](Visio.DrawingControl.DataRecordsetAdded.md)
- [DesignModeEntered](Visio.DrawingControl.DesignModeEntered.md)
- [DocumentChanged](Visio.DrawingControl.DocumentChanged.md)
- [DocumentCloseCanceled](Visio.DrawingControl.DocumentCloseCanceled.md)
- [DocumentCreated](Visio.DrawingControl.DocumentCreated.md)
- [DocumentOpened](Visio.DrawingControl.DocumentOpened.md)
- [DocumentSaved](Visio.DrawingControl.DocumentSaved.md)
- [DocumentSavedAs](Visio.DrawingControl.DocumentSavedAs.md)
- [GroupCanceled](Visio.DrawingControl.GroupCanceled.md)
- [KeyDown](Visio.DrawingControl.KeyDown.md)
- [KeyPress](Visio.DrawingControl.KeyPress.md)
- [KeyUp](Visio.DrawingControl.KeyUp.md)
- [MasterAdded](Visio.DrawingControl.MasterAdded.md)
- [MasterChanged](Visio.DrawingControl.MasterChanged.md)
- [MasterDeleteCanceled](Visio.DrawingControl.MasterDeleteCanceled.md)
- [MouseDown](Visio.DrawingControl.MouseDown.md)
- [MouseMove](Visio.DrawingControl.MouseMove.md)
- [MouseUp](Visio.DrawingControl.MouseUp.md)
- [OnKeystrokeMessageForAddon](Visio.DrawingControl.OnKeystrokeMessageForAddon.md)
- [PageAdded](Visio.DrawingControl.PageAdded.md)
- [PageChanged](Visio.DrawingControl.PageChanged.md)
- [PageDeleteCanceled](Visio.DrawingControl.PageDeleteCanceled.md)
- [QueryCancelConvertToGroup](Visio.DrawingControl.QueryCancelConvertToGroup.md)
- [QueryCancelDocumentClose](Visio.DrawingControl.QueryCancelDocumentClose.md)
- [QueryCancelGroup](Visio.DrawingControl.QueryCancelGroup.md)
- [QueryCancelMasterDelete](Visio.DrawingControl.QueryCancelMasterDelete.md)
- [QueryCancelPageDelete](Visio.DrawingControl.QueryCancelPageDelete.md)
- [QueryCancelSelectionDelete](Visio.DrawingControl.QueryCancelSelectionDelete.md)
- [QueryCancelStyleDelete](Visio.DrawingControl.QueryCancelStyleDelete.md)
- [QueryCancelUngroup](Visio.DrawingControl.QueryCancelUngroup.md)
- [QueryCancelWindowClose](Visio.DrawingControl.QueryCancelWindowClose.md)
- [RunModeEntered](Visio.DrawingControl.RunModeEntered.md)
- [SelectionChanged](Visio.DrawingControl.SelectionChanged.md)
- [SelectionDeleteCanceled](Visio.DrawingControl.SelectionDeleteCanceled.md)
- [ShapeAdded](Visio.DrawingControl.ShapeAdded.md)
- [ShapeDataGraphicChanged](Visio.DrawingControl.ShapeDataGraphicChanged.md)
- [ShapeExitedTextEdit](Visio.DrawingControl.ShapeExitedTextEdit.md)
- [ShapeParentChanged](Visio.DrawingControl.ShapeParentChanged.md)
- [StyleAdded](Visio.DrawingControl.StyleAdded.md)
- [StyleChanged](Visio.DrawingControl.StyleChanged.md)
- [StyleDeleteCanceled](Visio.DrawingControl.StyleDeleteCanceled.md)
- [UngroupCanceled](Visio.DrawingControl.UngroupCanceled.md)
- [ViewChanged](Visio.DrawingControl.ViewChanged.md)
- [WindowActivated](Visio.DrawingControl.WindowActivated.md)
- [WindowChanged](Visio.DrawingControl.WindowChanged.md)
- [WindowCloseCanceled](Visio.DrawingControl.WindowCloseCanceled.md)
- [WindowTurnedToPage](Visio.DrawingControl.WindowTurnedToPage.md)

## Properties

- [Document](Visio.DrawingControl.Document.md)
- [HostID](Visio.DrawingControl.HostID.md)
- [NegotiateMenus](Visio.DrawingControl.NegotiateMenus.md)
- [NegotiateToolbars](Visio.DrawingControl.NegotiateToolbars.md)
- [PageSizingBehavior](Visio.DrawingControl.PageSizingBehavior.md)
- [ShutDownBehavior](Visio.DrawingControl.ShutDownBehavior.md)
- [Src](Visio.DrawingControl.Src.md)
- [Window](Visio.DrawingControl.Window.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]