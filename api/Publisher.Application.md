---
title: Application object (Publisher)
keywords: vbapb10.chm536936447
f1_keywords:
- vbapb10.chm536936447
ms.prod: publisher
api_name:
- Publisher.Application
ms.assetid: acfc7efb-e6a5-a89a-3aee-3cb4af2f3508
ms.date: 05/31/2019
localization_priority: Normal
---


# Application object (Publisher)

Represents the Microsoft Publisher application. The **Application** object includes properties and methods that return top-level objects. For example, the **ActiveDocument** property returns a **Document** object.


## Remarks

When using Microsoft Visual Basic for Applications in Publisher, all the properties and methods of the **Application** object can be used without the **Application** object qualifier. For example, instead of typing **Application.ActiveDocument.PrintOut**, you can type **ActiveDocument.PrintOut**. 

Properties and methods that can be used without the **Application** object qualifier are considered _global_. To view the global properties and methods in the Object Browser, choose `<globals>` at the top of the list in the **Classes** box. 

When accessing the Publisher object model from a non-Publisher project, all properties and methods must be fully qualified.

Use the **[Application](Publisher.Application.Application.md)** property to return the **Application** object.


## Example

The following example displays the application name.

```vb
Sub ShowAppName() 
 MsgBox Application.Name 
End Sub
```


## Events

- [AfterPrint](Publisher.Application.AfterPrint.md)
- [BeforePrint](Publisher.Application.BeforePrint.md)
- [DocumentBeforeClose](Publisher.Application.DocumentBeforeClose.md)
- [DocumentOpen](Publisher.Application.DocumentOpen.md)
- [HideCatalogUI](Publisher.Application.HideCatalogUI.md)
- [MailMergeAfterMerge](Publisher.Application.MailMergeAfterMerge.md)
- [MailMergeAfterRecordMerge](Publisher.Application.MailMergeAfterRecordMerge.md)
- [MailMergeBeforeMerge](Publisher.Application.MailMergeBeforeMerge.md)
- [MailMergeBeforeRecordMerge](Publisher.Application.MailMergeBeforeRecordMerge.md)
- [MailMergeDataSourceLoad](Publisher.Application.MailMergeDataSourceLoad.md)
- [MailMergeDataSourceValidate](Publisher.Application.MailMergeDataSourceValidate.md)
- [MailMergeGenerateBarcode](Publisher.Application.MailMergeGenerateBarcode.md)
- [MailMergeInsertBarcode](Publisher.Application.MailMergeInsertBarcode.md)
- [MailMergeRecipientListClose](Publisher.Application.MailMergeRecipientListClose.md)
- [MailMergeWizardFollowUpCustom](Publisher.Application.MailMergeWizardFollowUpCustom.md)
- [MailMergeWizardStateChange](Publisher.Application.MailMergeWizardStateChange.md)
- [NewDocument](Publisher.Application.NewDocument(even).md)
- [Quit](Publisher.Application.Quit(even).md)
- [ShowCatalogUI](Publisher.Application.ShowCatalogUI.md)
- [WindowActivate](Publisher.Application.WindowActivate.md)
- [WindowDeactivate](Publisher.Application.WindowDeactivate.md)
- [WindowPageChange](Publisher.Application.WindowPageChange.md)

## Methods

- [CentimetersToPoints](Publisher.Application.CentimetersToPoints.md)
- [ChangeFileOpenDirectory](Publisher.Application.ChangeFileOpenDirectory.md)
- [EmusToPoints](Publisher.Application.EmusToPoints.md)
- [Help](Publisher.Application.Help.md)
- [InchesToPoints](Publisher.Application.InchesToPoints.md)
- [IsValidObject](Publisher.Application.IsValidObject.md)
- [LinesToPoints](Publisher.Application.LinesToPoints.md)
- [MillimetersToPoints](Publisher.Application.MillimetersToPoints.md)
- [NewDocument](Publisher.Application.NewDocument(method).md)
- [Open](Publisher.Application.Open.md)
- [PicasToPoints](Publisher.Application.PicasToPoints.md)
- [PixelsToPoints](Publisher.Application.PixelsToPoints.md)
- [PointsToCentimeters](Publisher.Application.PointsToCentimeters.md)
- [PointsToEmus](Publisher.Application.PointsToEmus.md)
- [PointsToInches](Publisher.Application.PointsToInches.md)
- [PointsToLines](Publisher.Application.PointsToLines.md)
- [PointsToMillimeters](Publisher.Application.PointsToMillimeters.md)
- [PointsToPicas](Publisher.Application.PointsToPicas.md)
- [PointsToPixels](Publisher.Application.PointsToPixels.md)
- [PointsToTwips](Publisher.Application.PointsToTwips.md)
- [Quit](Publisher.Application.Quit(method).md)
- [ShowWizardCatalog](Publisher.Application.ShowWizardCatalog.md)
- [TwipsToPoints](Publisher.Application.TwipsToPoints.md)

## Properties

- [ActiveDocument](Publisher.Application.ActiveDocument.md)
- [ActiveWindow](Publisher.Application.ActiveWindow.md)
- [Application](Publisher.Application.Application.md)
- [Assistance](Publisher.Application.Assistance.md)
- [AutomationSecurity](Publisher.Application.AutomationSecurity.md)
- [Build](Publisher.Application.Build.md)
- [CaptionStyles](Publisher.application.captionstyles.md)
- [ColorSchemes](Publisher.Application.ColorSchemes.md)
- [COMAddIns](Publisher.Application.COMAddIns.md)
- [CommandBars](Publisher.Application.CommandBars.md)
- [Documents](Publisher.Application.Documents.md)
- [FileDialog](Publisher.Application.FileDialog.md)
- [InsertBarcodeVisible](Publisher.Application.InsertBarcodeVisible.md)
- [InstalledPrinters](Publisher.Application.InstalledPrinters.md)
- [Language](Publisher.Application.Language.md)
- [Name](Publisher.Application.Name.md)
- [OfficeDataSourceObject](Publisher.Application.OfficeDataSourceObject.md)
- [Options](Publisher.Application.Options.md)
- [Parent](Publisher.Application.Parent.md)
- [Path](Publisher.Application.Path.md)
- [PathSeparator](Publisher.Application.PathSeparator.md)
- [PrintPreview](Publisher.Application.PrintPreview.md)
- [ProductCode](Publisher.Application.ProductCode.md)
- [ScreenUpdating](Publisher.Application.ScreenUpdating.md)
- [Selection](Publisher.Application.Selection.md)
- [ShowFollowUpCustom](Publisher.Application.ShowFollowUpCustom.md)
- [SnapToGuides](Publisher.Application.SnapToGuides.md)
- [SnapToObjects](Publisher.Application.SnapToObjects.md)
- [TemplateFolderPath](Publisher.Application.TemplateFolderPath.md)
- [ValidateAddressVisible](Publisher.Application.ValidateAddressVisible.md)
- [Version](Publisher.Application.Version.md)
- [WebOptions](Publisher.Application.WebOptions.md)
- [WizardCatalogVisible](Publisher.Application.WizardCatalogVisible.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]