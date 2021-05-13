---
title: Presentation object (PowerPoint)
description: Represents a Microsoft PowerPoint presentation. 
keywords: vbapp10.chm524000
f1_keywords:
- vbapp10.chm524000
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation
ms.assetid: ec75cf52-69f8-d35b-0a26-4a8da8a9683f
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation object (PowerPoint)

Represents a Microsoft PowerPoint presentation. 


## Remarks

The **Presentation** object is a member of the **[Presentations](PowerPoint.Presentations.md)** collection. The **Presentations** collection contains all the **Presentation** objects that represent open presentations in PowerPoint.

The following examples describe how to:


- Return a presentation that you specify by name or index number
    
- Return the presentation in the active window
    
- Return the presentation in any document window or slide show window you specify
    

## Example

Use  **Presentations** (_index_), where _index_ is the presentation's name or index number, to return a single **Presentation** object. The name of the presentation is the file name, with or without the file name extension, and without the path. The following example adds a slide to the beginning of Sample Presentation.


```vb
Presentations("Sample Presentation").Slides.Add 1, 1
```

Note that if multiple presentations with the same name are open, the first presentation in the collection with the specified name is returned.

Use the [ActivePresentation](PowerPoint.Application.ActivePresentation.md) property to return the presentation in the active window. The following example saves the active presentation.




```vb
ActivePresentation.Save
```

Use the [Presentation](PowerPoint.DocumentWindow.Presentation.md) property to return the presentation that's in the specified document window or slide show window. The following example displays the name of the slide show running in slide show window one.




```vb
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|Name|
|:-----|
|[AcceptAll](PowerPoint.Presentation.AcceptAll.md)|
|[AddTitleMaster](PowerPoint.Presentation.AddTitleMaster.md)|
|[AddToFavorites](PowerPoint.Presentation.AddToFavorites.md)|
|[ApplyTemplate](PowerPoint.Presentation.ApplyTemplate.md)|
|[ApplyTemplate2](PowerPoint.presentation.applytemplate2.md)|
|[ApplyTheme](PowerPoint.Presentation.ApplyTheme.md)|
|[CanCheckIn](PowerPoint.Presentation.CanCheckIn.md)|
|[CheckIn](PowerPoint.Presentation.CheckIn.md)|
|[CheckInWithVersion](PowerPoint.Presentation.CheckInWithVersion.md)|
|[Close](PowerPoint.Presentation.Close.md)|
|[Convert2](PowerPoint.Presentation.Convert2.md)|
|[CreateVideo](PowerPoint.Presentation.CreateVideo.md)|
|[EndReview](PowerPoint.Presentation.EndReview.md)|
|[EnsureAllMediaUpgraded](PowerPoint.Presentation.EnsureAllMediaUpgraded.md)|
|[Export](PowerPoint.Presentation.Export.md)|
|[ExportAsFixedFormat](PowerPoint.Presentation.ExportAsFixedFormat.md)|
|[ExportAsFixedFormat2](PowerPoint.presentation.exportasfixedformat2.md)|
|[FollowHyperlink](PowerPoint.Presentation.FollowHyperlink.md)|
|[GetWorkflowTasks](PowerPoint.Presentation.GetWorkflowTasks.md)|
|[GetWorkflowTemplates](PowerPoint.Presentation.GetWorkflowTemplates.md)|
|[LockServerFile](PowerPoint.Presentation.LockServerFile.md)|
|[Merge](PowerPoint.presentation.merge.md)|
|[MergeWithBaseline](PowerPoint.Presentation.MergeWithBaseline.md)|
|[NewWindow](PowerPoint.Presentation.NewWindow.md)|
|[PrintOut](PowerPoint.Presentation.PrintOut.md)|
|[PublishSlides](PowerPoint.Presentation.PublishSlides.md)|
|[RejectAll](PowerPoint.Presentation.RejectAll.md)|
|[RemoveDocumentInformation](PowerPoint.Presentation.RemoveDocumentInformation.md)|
|[Save](PowerPoint.Presentation.Save.md)|
|[SaveAs](PowerPoint.Presentation.SaveAs.md)|
|[SaveCopyAs](PowerPoint.Presentation.SaveCopyAs.md)|
|[SaveCopyAs2](PowerPoint.Presentation.SaveCopyAs2.md)|
|[SendFaxOverInternet](PowerPoint.Presentation.SendFaxOverInternet.md)|
|[SetPasswordEncryptionOptions](PowerPoint.Presentation.SetPasswordEncryptionOptions.md)|
|[UpdateLinks](PowerPoint.Presentation.UpdateLinks.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Presentation.Application.md)|
|[Broadcast](PowerPoint.Presentation.Broadcast.md)|
|[BuiltInDocumentProperties](PowerPoint.Presentation.BuiltInDocumentProperties.md)|
|[ChartDataPointTrack](PowerPoint.presentation.chartdatapointtrack.md)|
|[Coauthoring](PowerPoint.Presentation.Coauthoring.md)|
|[ColorSchemes](PowerPoint.Presentation.ColorSchemes.md)|
|[CommandBars](PowerPoint.Presentation.CommandBars.md)|
|[Container](PowerPoint.Presentation.Container.md)|
|[ContentTypeProperties](PowerPoint.Presentation.ContentTypeProperties.md)|
|[CreateVideoStatus](PowerPoint.Presentation.CreateVideoStatus.md)|
|[CustomDocumentProperties](PowerPoint.Presentation.CustomDocumentProperties.md)|
|[CustomerData](PowerPoint.Presentation.CustomerData.md)|
|[CustomXMLParts](PowerPoint.Presentation.CustomXMLParts.md)|
|[DefaultLanguageID](PowerPoint.Presentation.DefaultLanguageID.md)|
|[DefaultShape](PowerPoint.Presentation.DefaultShape.md)|
|[Designs](PowerPoint.Presentation.Designs.md)|
|[DisplayComments](PowerPoint.Presentation.DisplayComments.md)|
|[DocumentInspectors](PowerPoint.Presentation.DocumentInspectors.md)|
|[DocumentLibraryVersions](PowerPoint.Presentation.DocumentLibraryVersions.md)|
|[EncryptionProvider](PowerPoint.Presentation.EncryptionProvider.md)|
|[EnvelopeVisible](PowerPoint.Presentation.EnvelopeVisible.md)|
|[ExtraColors](PowerPoint.Presentation.ExtraColors.md)|
|[FarEastLineBreakLanguage](PowerPoint.Presentation.FarEastLineBreakLanguage.md)|
|[FarEastLineBreakLevel](PowerPoint.Presentation.FarEastLineBreakLevel.md)|
|[Final](PowerPoint.Presentation.Final.md)|
|[Fonts](PowerPoint.Presentation.Fonts.md)|
|[FullName](PowerPoint.Presentation.FullName.md)|
|[GridDistance](PowerPoint.Presentation.GridDistance.md)|
|[Guides](PowerPoint.presentation.guides.md)|
|[HandoutMaster](PowerPoint.Presentation.HandoutMaster.md)|
|[HasHandoutMaster](PowerPoint.Presentation.HasHandoutMaster.md)|
|[HasNotesMaster](PowerPoint.Presentation.HasNotesMaster.md)|
|[HasTitleMaster](PowerPoint.Presentation.HasTitleMaster.md)|
|[HasVBProject](PowerPoint.Presentation.HasVBProject.md)|
|[InMergeMode](PowerPoint.Presentation.InMergeMode.md)|
|[LayoutDirection](PowerPoint.Presentation.LayoutDirection.md)|
|[Name](PowerPoint.Presentation.Name.md)|
|[NoLineBreakAfter](PowerPoint.Presentation.NoLineBreakAfter.md)|
|[NoLineBreakBefore](PowerPoint.Presentation.NoLineBreakBefore.md)|
|[NotesMaster](PowerPoint.Presentation.NotesMaster.md)|
|[PageSetup](PowerPoint.Presentation.PageSetup.md)|
|[Parent](PowerPoint.Presentation.Parent.md)|
|[Password](PowerPoint.Presentation.Password.md)|
|[PasswordEncryptionAlgorithm](PowerPoint.Presentation.PasswordEncryptionAlgorithm.md)|
|[PasswordEncryptionFileProperties](PowerPoint.Presentation.PasswordEncryptionFileProperties.md)|
|[PasswordEncryptionKeyLength](PowerPoint.Presentation.PasswordEncryptionKeyLength.md)|
|[PasswordEncryptionProvider](PowerPoint.Presentation.PasswordEncryptionProvider.md)|
|[Path](PowerPoint.Presentation.Path.md)|
|[Permission](PowerPoint.Presentation.Permission.md)|
|[PrintOptions](PowerPoint.Presentation.PrintOptions.md)|
|[ReadOnly](PowerPoint.Presentation.ReadOnly.md)|
|[ReadOnlyRecommended](PowerPoint.Presentation.ReadOnlyRecommended.md)|
|[RemovePersonalInformation](PowerPoint.Presentation.RemovePersonalInformation.md)|
|[Research](PowerPoint.Presentation.Research.md)|
|[Saved](PowerPoint.Presentation.Saved.md)|
|[SectionProperties](PowerPoint.Presentation.SectionProperties.md)|
|[ServerPolicy](PowerPoint.Presentation.ServerPolicy.md)|
|[SharedWorkspace](PowerPoint.Presentation.SharedWorkspace.md)|
|[Signatures](PowerPoint.Presentation.Signatures.md)|
|[SlideMaster](PowerPoint.Presentation.SlideMaster.md)|
|[Slides](PowerPoint.Presentation.Slides.md)|
|[SlideShowSettings](PowerPoint.Presentation.SlideShowSettings.md)|
|[SlideShowWindow](PowerPoint.Presentation.SlideShowWindow.md)|
|[SnapToGrid](PowerPoint.Presentation.SnapToGrid.md)|
|[Sync](PowerPoint.Presentation.Sync.md)|
|[Tags](PowerPoint.Presentation.Tags.md)|
|[TemplateName](PowerPoint.Presentation.TemplateName.md)|
|[TitleMaster](PowerPoint.Presentation.TitleMaster.md)|
|[VBASigned](PowerPoint.Presentation.VBASigned.md)|
|[VBProject](PowerPoint.Presentation.VBProject.md)|
|[Windows](PowerPoint.Presentation.Windows.md)|
|[WritePassword](PowerPoint.Presentation.WritePassword.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
