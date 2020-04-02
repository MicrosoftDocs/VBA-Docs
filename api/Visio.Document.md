---
title: Document object (Visio)
keywords: vis_sdr.chm10080
f1_keywords:
- vis_sdr.chm10080
ms.prod: visio
api_name:
- Visio.Document
ms.assetid: 21640062-13a2-a2b2-7c61-7e707671207c
ms.date: 06/19/2019
localization_priority: Normal
---


# Document object (Visio)

Represents a drawing file (.vsd or .vdx), stencil file (.vss or .vsx), or template file (.vst or .vtx) that is open in an instance of Microsoft Visio. A **Document** object is a member of the **[Documents](Visio.Documents.md)** collection of an **Application** object.


## Remarks

The default property of a **Document** object is **Name**.

Use the **[Open](visio.documents.open.md)** method of a **Documents** collection to open an existing document.

Use the **[Add](visio.documents.add.md)** method of a **Documents** collection to create a new document.

Use the **[ActiveDocument](visio.application.activedocument.md)** property of an **Application** object to retrieve the active document in an instance.

Use the **Pages**, **Masters**, and **Styles** properties to retrieve **Page**, **Master**, and **Style** objects, respectively.

Use the **CustomMenus** or **CustomToolbars** properties to access the custom menus or toolbars.

> [!NOTE] 
> The Microsoft Visual Basic for Applications (VBA) project of every Visio document also has a class module called **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)**. When you reference the **ThisDocument** module from code in a VBA project, it returns a reference to the project's **Document** object. For example, the code in a document's project can display the name of the project's document in a **message** box with this statement:
> 
> ```vb
>    MsgBox ThisDocument.Name
> ```

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this object maps to the following types:

- **Microsoft.Office.Interop.Visio.IVDocument**
    

## Events

- [AfterDocumentMerge](Visio.document.afterdocumentmerge.md)
- [AfterRemoveHiddenInformation](Visio.Document.AfterRemoveHiddenInformation.md)
- [BeforeDataRecordsetDelete](Visio.Document.BeforeDataRecordsetDelete.md)
- [BeforeDocumentClose](Visio.Document.BeforeDocumentClose.md)
- [BeforeDocumentSave](Visio.Document.BeforeDocumentSave.md)
- [BeforeDocumentSaveAs](Visio.Document.BeforeDocumentSaveAs.md)
- [BeforeMasterDelete](Visio.Document.BeforeMasterDelete.md)
- [BeforePageDelete](Visio.Document.BeforePageDelete.md)
- [BeforeSelectionDelete](Visio.Document.BeforeSelectionDelete.md)
- [BeforeShapeTextEdit](Visio.Document.BeforeShapeTextEdit.md)
- [BeforeStyleDelete](Visio.Document.BeforeStyleDelete.md)
- [ConvertToGroupCanceled](Visio.Document.ConvertToGroupCanceled.md)
- [DataRecordsetAdded](Visio.Document.DataRecordsetAdded.md)
- [DesignModeEntered](Visio.Document.DesignModeEntered.md)
- [DocumentChanged](Visio.Document.DocumentChanged.md)
- [DocumentCloseCanceled](Visio.Document.DocumentCloseCanceled.md)
- [DocumentCreated](Visio.Document.DocumentCreated.md)
- [DocumentOpened](Visio.Document.DocumentOpened.md)
- [DocumentSaved](Visio.Document.DocumentSaved.md)
- [DocumentSavedAs](Visio.Document.DocumentSavedAs.md)
- [GroupCanceled](Visio.Document.GroupCanceled.md)
- [MasterAdded](Visio.Document.MasterAdded.md)
- [MasterChanged](Visio.Document.MasterChanged.md)
- [MasterDeleteCanceled](Visio.Document.MasterDeleteCanceled.md)
- [PageAdded](Visio.Document.PageAdded.md)
- [PageChanged](Visio.Document.PageChanged.md)
- [PageDeleteCanceled](Visio.Document.PageDeleteCanceled.md)
- [QueryCancelConvertToGroup](Visio.Document.QueryCancelConvertToGroup.md)
- [QueryCancelDocumentClose](Visio.Document.QueryCancelDocumentClose.md)
- [QueryCancelGroup](Visio.Document.QueryCancelGroup.md)
- [QueryCancelMasterDelete](Visio.Document.QueryCancelMasterDelete.md)
- [QueryCancelPageDelete](Visio.Document.QueryCancelPageDelete.md)
- [QueryCancelSelectionDelete](Visio.Document.QueryCancelSelectionDelete.md)
- [QueryCancelStyleDelete](Visio.Document.QueryCancelStyleDelete.md)
- [QueryCancelUngroup](Visio.Document.QueryCancelUngroup.md)
- [RuleSetValidated](Visio.Document.RuleSetValidated.md)
- [RunModeEntered](Visio.Document.RunModeEntered.md)
- [SelectionDeleteCanceled](Visio.Document.SelectionDeleteCanceled.md)
- [ShapeAdded](Visio.Document.ShapeAdded.md)
- [ShapeDataGraphicChanged](Visio.Document.ShapeDataGraphicChanged.md)
- [ShapeExitedTextEdit](Visio.Document.ShapeExitedTextEdit.md)
- [ShapeParentChanged](Visio.Document.ShapeParentChanged.md)
- [StyleAdded](Visio.Document.StyleAdded.md)
- [StyleChanged](Visio.Document.StyleChanged.md)
- [StyleDeleteCanceled](Visio.Document.StyleDeleteCanceled.md)
- [UngroupCanceled](Visio.Document.UngroupCanceled.md)

## Methods

- [AddUndoUnit](Visio.Document.AddUndoUnit.md)
- [BeginUndoScope](Visio.Document.BeginUndoScope.md)
- [CanCheckIn](Visio.Document.CanCheckIn.md)
- [CanUndoCheckOut](Visio.Document.CanUndoCheckOut.md)
- [CheckIn](Visio.Document.CheckIn.md)
- [Clean](Visio.Document.Clean.md)
- [ClearCustomMenus](Visio.Document.ClearCustomMenus.md)
- [ClearCustomToolbars](Visio.Document.ClearCustomToolbars.md)
- [ClearGestureFormatSheet](Visio.Document.ClearGestureFormatSheet.md)
- [Close](Visio.Document.Close.md)
- [CopyPreviewPicture](Visio.Document.CopyPreviewPicture.md)
- [DeleteSolutionXMLElement](Visio.Document.DeleteSolutionXMLElement.md)
- [Drop](Visio.Document.Drop.md)
- [EndUndoScope](Visio.Document.EndUndoScope.md)
- [ExecuteLine](Visio.Document.ExecuteLine.md)
- [ExportAsFixedFormat](Visio.Document.ExportAsFixedFormat.md)
- [FollowHyperlink](Visio.Document.FollowHyperlink.md)
- [GetThemeNames](Visio.Document.GetThemeNames.md)
- [GetThemeNamesU](Visio.Document.GetThemeNamesU.md)
- [OpenStencilWindow](Visio.Document.OpenStencilWindow.md)
- [ParseLine](Visio.Document.ParseLine.md)
- [Print](Visio.Document.Print.md)
- [PrintOut](Visio.Document.PrintOut.md)
- [PurgeUndo](Visio.Document.PurgeUndo.md)
- [RemoveHiddenInformation](Visio.Document.RemoveHiddenInformation.md)
- [RenameCurrentScope](Visio.Document.RenameCurrentScope.md)
- [Save](Visio.Document.Save.md)
- [SaveAs](Visio.Document.SaveAs.md)
- [SaveAsEx](Visio.Document.SaveAsEx.md)
- [SetCustomMenus](Visio.Document.SetCustomMenus.md)
- [SetCustomToolbars](Visio.Document.SetCustomToolbars.md)
- [UndoCheckOut](Visio.Document.UndoCheckOut.md)

## Properties

- [AlternateNames](Visio.Document.AlternateNames.md)
- [Application](Visio.Document.Application.md)
- [AutoRecover](Visio.Document.AutoRecover.md)
- [BottomMargin](Visio.Document.BottomMargin.md)
- [BuildNumberCreated](Visio.Document.BuildNumberCreated.md)
- [BuildNumberEdited](Visio.Document.BuildNumberEdited.md)
- [Category](Visio.Document.Category.md)
- [ClassID](Visio.Document.ClassID.md)
- [Colors](Visio.Document.Colors.md)
- [Comments](Visio.document.comments.md)
- [Company](Visio.Document.Company.md)
- [CompatibilityMode](Visio.document.compatibilitymode.md)
- [Container](Visio.Document.Container.md)
- [ContainsWorkspaceEx](Visio.Document.ContainsWorkspaceEx.md)
- [Creator](Visio.Document.Creator.md)
- [CustomMenus](Visio.Document.CustomMenus.md)
- [CustomMenusFile](Visio.Document.CustomMenusFile.md)
- [CustomToolbars](Visio.Document.CustomToolbars.md)
- [CustomToolbarsFile](Visio.Document.CustomToolbarsFile.md)
- [CustomUI](Visio.Document.CustomUI.md)
- [DataRecordsets](Visio.Document.DataRecordsets.md)
- [DefaultFillStyle](Visio.Document.DefaultFillStyle.md)
- [DefaultGuideStyle](Visio.Document.DefaultGuideStyle.md)
- [DefaultLineStyle](Visio.Document.DefaultLineStyle.md)
- [DefaultSavePath](Visio.Document.DefaultSavePath.md)
- [DefaultStyle](Visio.Document.DefaultStyle.md)
- [DefaultTextStyle](Visio.Document.DefaultTextStyle.md)
- [Description](Visio.Document.Description.md)
- [DiagramServicesEnabled](Visio.Document.DiagramServicesEnabled.md)
- [DocumentSheet](Visio.Document.DocumentSheet.md)
- [DynamicGridEnabled](Visio.Document.DynamicGridEnabled.md)
- [EditorCount](Visio.document.editorcount.md)
- [EmailRoutingData](Visio.Document.EmailRoutingData.md)
- [EventList](Visio.Document.EventList.md)
- [Fonts](Visio.Document.Fonts.md)
- [FooterCenter](Visio.Document.FooterCenter.md)
- [FooterLeft](Visio.Document.FooterLeft.md)
- [FooterMargin](Visio.Document.FooterMargin.md)
- [FooterRight](Visio.Document.FooterRight.md)
- [FullBuildNumberCreated](Visio.Document.FullBuildNumberCreated.md)
- [FullBuildNumberEdited](Visio.Document.FullBuildNumberEdited.md)
- [FullName](Visio.Document.FullName.md)
- [GestureFormatSheet](Visio.Document.GestureFormatSheet.md)
- [GlueEnabled](Visio.Document.GlueEnabled.md)
- [GlueSettings](Visio.Document.GlueSettings.md)
- [HeaderCenter](Visio.Document.HeaderCenter.md)
- [HeaderFooterColor](Visio.Document.HeaderFooterColor.md)
- [HeaderFooterFont](Visio.Document.HeaderFooterFont.md)
- [HeaderLeft](Visio.Document.HeaderLeft.md)
- [HeaderMargin](Visio.Document.HeaderMargin.md)
- [HeaderRight](Visio.Document.HeaderRight.md)
- [HyperlinkBase](Visio.Document.HyperlinkBase.md)
- [ID](Visio.Document.ID.md)
- [Index](Visio.Document.Index.md)
- [InPlace](Visio.Document.InPlace.md)
- [Keywords](Visio.Document.Keywords.md)
- [Language](Visio.Document.Language.md)
- [LeftMargin](Visio.Document.LeftMargin.md)
- [MacrosEnabled](Visio.Document.MacrosEnabled.md)
- [Manager](Visio.Document.Manager.md)
- [Masters](Visio.Document.Masters.md)
- [MasterShortcuts](Visio.Document.MasterShortcuts.md)
- [Mode](Visio.Document.Mode.md)
- [Name](Visio.Document.Name.md)
- [ObjectType](Visio.Document.ObjectType.md)
- [OLEObjects](Visio.Document.OLEObjects.md)
- [Pages](Visio.Document.Pages.md)
- [PaperHeight](Visio.Document.PaperHeight.md)
- [PaperSize](Visio.Document.PaperSize.md)
- [PaperWidth](Visio.Document.PaperWidth.md)
- [Path](Visio.Document.Path.md)
- [Permission](Visio.document.permission.md)
- [PersistsEvents](Visio.Document.PersistsEvents.md)
- [PreviewPicture](Visio.Document.PreviewPicture.md)
- [PrintCenteredH](Visio.Document.PrintCenteredH.md)
- [PrintCenteredV](Visio.Document.PrintCenteredV.md)
- [Printer](Visio.Document.Printer.md)
- [PrintFitOnPages](Visio.Document.PrintFitOnPages.md)
- [PrintLandscape](Visio.Document.PrintLandscape.md)
- [PrintPagesAcross](Visio.Document.PrintPagesAcross.md)
- [PrintPagesDown](Visio.Document.PrintPagesDown.md)
- [PrintScale](Visio.Document.PrintScale.md)
- [ProgID](Visio.Document.ProgID.md)
- [Protection](Visio.Document.Protection.md)
- [ReadOnly](Visio.Document.ReadOnly.md)
- [RemovePersonalInformation](Visio.Document.RemovePersonalInformation.md)
- [RightMargin](Visio.Document.RightMargin.md)
- [Saved](Visio.Document.Saved.md)
- [SavePreviewMode](Visio.Document.SavePreviewMode.md)
- [ServerPublishOptions](Visio.Document.ServerPublishOptions.md)
- [SharedWorkspace](Visio.Document.SharedWorkspace.md)
- [SnapAngles](Visio.Document.SnapAngles.md)
- [SnapEnabled](Visio.Document.SnapEnabled.md)
- [SnapExtensions](Visio.Document.SnapExtensions.md)
- [SnapSettings](Visio.Document.SnapSettings.md)
- [SolutionXMLElement](Visio.Document.SolutionXMLElement.md)
- [SolutionXMLElementCount](Visio.Document.SolutionXMLElementCount.md)
- [SolutionXMLElementExists](Visio.Document.SolutionXMLElementExists.md)
- [SolutionXMLElementName](Visio.Document.SolutionXMLElementName.md)
- [Stat](Visio.Document.Stat.md)
- [Styles](Visio.Document.Styles.md)
- [Subject](Visio.Document.Subject.md)
- [Sync](Visio.Document.Sync.md)
- [Template](Visio.Document.Template.md)
- [Time](Visio.Document.Time.md)
- [TimeCreated](Visio.Document.TimeCreated.md)
- [TimeEdited](Visio.Document.TimeEdited.md)
- [TimePrinted](Visio.Document.TimePrinted.md)
- [TimeSaved](Visio.Document.TimeSaved.md)
- [Title](Visio.Document.Title.md)
- [TopMargin](Visio.Document.TopMargin.md)
- [Type](Visio.Document.Type.md)
- [UndoEnabled](Visio.Document.UndoEnabled.md)
- [UserCustomUI](Visio.Document.UserCustomUI.md)
- [Validation](Visio.Document.Validation.md)
- [VBProject](Visio.Document.VBProject.md)
- [VBProjectData](Visio.Document.VBProjectData.md)
- [Version](Visio.Document.Version.md)
- [ZoomBehavior](Visio.Document.ZoomBehavior.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]