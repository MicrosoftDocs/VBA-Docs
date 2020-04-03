---
title: Application object (PowerPoint)
keywords: vbapp10.chm504000
f1_keywords:
- vbapp10.chm504000
ms.prod: powerpoint
api_name:
- PowerPoint.Application
ms.assetid: 978c2b99-4271-b953-4283-73b5f3d96f41
ms.date: 06/08/2017
localization_priority: Normal
---


# Application object (PowerPoint)

Represents the entire Microsoft PowerPoint application. 


## Remarks

The  **Application** object contains:


- Application-wide settings and options (the name of the active printer, for example).
    
- Properties that return top-level objects, such as  **ActivePresentation**, and **Windows**.
    


When you are writing code that will run from PowerPoint, you can use the following properties of the  **Application** object without the object qualifier: **ActivePresentation**, **ActiveWindow**, **AddIns**, **Presentations**, **SlideShowWindows**, **Windows**.

For example, instead of writing  `Application.ActiveWindow.Height = 200`, you can write  `ActiveWindow.Height = 200`.


## Example

Use the  **Application** property to return the **Application** object. The following example returns the path to the program file.


```vb
Dim MyPath As String

MyPath = Application.Path
```

The following example creates a PowerPoint  **Application** object in another application, starts PowerPoint (if it is not already running), and opens an existing presentation named "Ex_a2a.ppt".




```vb
Set ppt = New Powerpoint.Application

ppt.Visible = True

ppt.Presentations.Open "c:\My Documents\ex_a2a.ppt"
```


## Events



|Name|
|:-----|
|[AfterDragDropOnSlide](PowerPoint.application.afterdragdroponslide.md)|
|[AfterNewPresentation](PowerPoint.Application.AfterNewPresentation.md)|
|[AfterPresentationOpen](PowerPoint.Application.AfterPresentationOpen.md)|
|[AfterShapeSizeChange](PowerPoint.application.aftershapesizechange.md)|
|[ColorSchemeChanged](PowerPoint.Application.ColorSchemeChanged.md)|
|[NewPresentation](PowerPoint.Application.NewPresentation(even).md)|
|[PresentationBeforeClose](PowerPoint.Application.PresentationBeforeClose.md)|
|[PresentationBeforeSave](PowerPoint.Application.PresentationBeforeSave.md)|
|[PresentationClose](PowerPoint.Application.PresentationClose.md)|
|[PresentationCloseFinal](PowerPoint.Application.PresentationCloseFinal.md)|
|[PresentationNewSlide](PowerPoint.Application.PresentationNewSlide.md)|
|[PresentationOpen](PowerPoint.Application.PresentationOpen.md)|
|[PresentationPrint](PowerPoint.Application.PresentationPrint.md)|
|[PresentationSave](PowerPoint.Application.PresentationSave.md)|
|[PresentationSync](PowerPoint.Application.PresentationSync.md)|
|[ProtectedViewWindowActivate](PowerPoint.Application.ProtectedViewWindowActivate.md)|
|[ProtectedViewWindowBeforeClose](PowerPoint.Application.ProtectedViewWindowBeforeClose.md)|
|[ProtectedViewWindowBeforeEdit](PowerPoint.Application.ProtectedViewWindowBeforeEdit.md)|
|[ProtectedViewWindowDeactivate](PowerPoint.Application.ProtectedViewWindowDeactivate.md)|
|[ProtectedViewWindowOpen](PowerPoint.Application.ProtectedViewWindowOpen.md)|
|[SlideSelectionChanged](PowerPoint.Application.SlideSelectionChanged.md)|
|[SlideShowBegin](PowerPoint.Application.SlideShowBegin.md)|
|[SlideShowEnd](PowerPoint.Application.SlideShowEnd.md)|
|[SlideShowNextBuild](PowerPoint.Application.SlideShowNextBuild.md)|
|[SlideShowNextClick](PowerPoint.Application.SlideShowNextClick.md)|
|[SlideShowNextSlide](PowerPoint.Application.SlideShowNextSlide.md)|
|[SlideShowOnNext](PowerPoint.Application.SlideShowOnNext.md)|
|[SlideShowOnPrevious](PowerPoint.Application.SlideShowOnPrevious.md)|
|[WindowActivate](PowerPoint.Application.WindowActivate.md)|
|[WindowBeforeDoubleClick](PowerPoint.Application.WindowBeforeDoubleClick.md)|
|[WindowBeforeRightClick](PowerPoint.Application.WindowBeforeRightClick.md)|
|[WindowDeactivate](PowerPoint.Application.WindowDeactivate.md)|
|[WindowSelectionChange](PowerPoint.Application.WindowSelectionChange.md)|

## Methods



|Name|
|:-----|
|[Activate](PowerPoint.Application.Activate.md)|
|[Help](PowerPoint.Application.Help.md)|
|[OpenThemeFile](PowerPoint.application.openthemefile.md)|
|[Quit](PowerPoint.Application.Quit.md)|
|[Run](PowerPoint.Application.Run.md)|
|[StartNewUndoEntry](PowerPoint.Application.StartNewUndoEntry.md)|

## Properties



|Name|
|:-----|
|[Active](PowerPoint.Application.Active.md)|
|[ActiveEncryptionSession](PowerPoint.Application.ActiveEncryptionSession.md)|
|[ActivePresentation](PowerPoint.Application.ActivePresentation.md)|
|[ActivePrinter](PowerPoint.Application.ActivePrinter.md)|
|[ActiveProtectedViewWindow](PowerPoint.Application.ActiveProtectedViewWindow.md)|
|[ActiveWindow](PowerPoint.Application.ActiveWindow.md)|
|[AddIns](PowerPoint.Application.AddIns.md)|
|[Assistance](PowerPoint.Application.Assistance.md)|
|[AutoCorrect](PowerPoint.Application.AutoCorrect.md)|
|[AutomationSecurity](PowerPoint.Application.AutomationSecurity.md)|
|[Build](PowerPoint.Application.Build.md)|
|[Caption](PowerPoint.Application.Caption.md)|
|[ChartDataPointTrack](PowerPoint.application.chartdatapointtrack.md)|
|[COMAddIns](PowerPoint.Application.COMAddIns.md)|
|[CommandBars](PowerPoint.Application.CommandBars.md)|
|[Creator](PowerPoint.Application.Creator.md)|
|[DisplayAlerts](PowerPoint.Application.DisplayAlerts.md)|
|[DisplayDocumentInformationPanel](PowerPoint.Application.DisplayDocumentInformationPanel.md)|
|[DisplayGridLines](PowerPoint.Application.DisplayGridLines.md)|
|[DisplayGuides](PowerPoint.application.displayguides.md)|
|[FeatureInstall](PowerPoint.Application.FeatureInstall.md)|
|[FileConverters](PowerPoint.Application.FileConverters.md)|
|[FileDialog](PowerPoint.Application.FileDialog.md)|
|[FileValidation](PowerPoint.Application.FileValidation.md)|
|[Height](PowerPoint.Application.Height.md)|
|[IsSandboxed](PowerPoint.Application.IsSandboxed.md)|
|[LanguageSettings](PowerPoint.Application.LanguageSettings.md)|
|[Left](PowerPoint.Application.Left.md)|
|[Name](PowerPoint.Application.Name.md)|
|[NewPresentation](PowerPoint.Application.NewPresentation(property).md)|
|[OperatingSystem](PowerPoint.Application.OperatingSystem.md)|
|[Options](PowerPoint.Application.Options.md)|
|[Path](PowerPoint.Application.Path.md)|
|[Presentations](PowerPoint.Application.Presentations.md)|
|[ProductCode](PowerPoint.Application.ProductCode.md)|
|[ProtectedViewWindows](PowerPoint.Application.ProtectedViewWindows.md)|
|[ShowStartupDialog](PowerPoint.Application.ShowStartupDialog.md)|
|[ShowWindowsInTaskbar](PowerPoint.Application.ShowWindowsInTaskbar.md)|
|[SlideShowWindows](PowerPoint.Application.SlideShowWindows.md)|
|[SmartArtColors](PowerPoint.Application.SmartArtColors.md)|
|[SmartArtLayouts](PowerPoint.Application.SmartArtLayouts.md)|
|[SmartArtQuickStyles](PowerPoint.Application.SmartArtQuickStyles.md)|
|[Top](PowerPoint.Application.Top.md)|
|[VBE](PowerPoint.Application.VBE.md)|
|[Version](PowerPoint.Application.Version.md)|
|[Visible](PowerPoint.Application.Visible.md)|
|[Width](PowerPoint.Application.Width.md)|
|[Windows](PowerPoint.Application.Windows.md)|
|[WindowState](PowerPoint.Application.WindowState.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
