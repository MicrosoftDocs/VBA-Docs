---
title: Options object (Publisher)
keywords: vbapb10.chm1114111
f1_keywords:
- vbapb10.chm1114111
ms.prod: publisher
api_name:
- Publisher.Options
ms.assetid: 2554cd33-9d94-2622-6fab-19ca33d5a561
ms.date: 06/01/2019
localization_priority: Normal
---


# Options object (Publisher)

Represents application and publication options in Microsoft Publisher. Many of the properties for the **Options** object correspond to items in the **Options** dialog box (**Tools** menu).

## Remarks

Use the **[Application.Options](Publisher.Application.Options.md)** property to return the **Options** object. 


## Example

The following example sets four application options for Publisher.

```vb
Sub SetSpecialOptions() 
 With Options 
 .AllowBackgroundSave = True 
 .DragAndDropText = True 
 .AutoHyphenate = True 
 .MeasurementUnit = pbUnitInch 
 End With 
End Sub
```


## Methods

- [ResetTips](Publisher.Options.ResetTips.md)
- [ResetWizardSynchronizing](Publisher.Options.ResetWizardSynchronizing.md)

## Properties

- [AddHebDoubleQuote](Publisher.Options.AddHebDoubleQuote.md)
- [AllowBackgroundSave](Publisher.Options.AllowBackgroundSave.md)
- [Application](Publisher.Options.Application.md)
- [AutoFormatWord](Publisher.Options.AutoFormatWord.md)
- [AutoHyphenate](Publisher.Options.AutoHyphenate.md)
- [AutoKeyboardSwitching](Publisher.Options.AutoKeyboardSwitching.md)
- [AutoSelectWord](Publisher.Options.AutoSelectWord.md)
- [DefaultPubDirection](Publisher.Options.DefaultPubDirection.md)
- [DefaultTextFlowDirection](Publisher.Options.DefaultTextFlowDirection.md)
- [DisplayStatusBar](Publisher.Options.DisplayStatusBar.md)
- [DragAndDropText](Publisher.Options.DragAndDropText.md)
- [HyphenationZone](Publisher.Options.HyphenationZone.md)
- [MeasurementUnit](Publisher.Options.MeasurementUnit.md)
- [Parent](Publisher.Options.Parent.md)
- [PathForPictures](Publisher.Options.PathForPictures.md)
- [PathForPublications](Publisher.Options.PathForPublications.md)
- [SaveAutoRecoverInfo](Publisher.Options.SaveAutoRecoverInfo.md)
- [SaveAutoRecoverInfoInterval](Publisher.Options.SaveAutoRecoverInfoInterval.md)
- [SequenceCheck](Publisher.Options.SequenceCheck.md)
- [ShowBasicColors](Publisher.Options.ShowBasicColors.md)
- [ShowScreenTipsOnObjects](Publisher.Options.ShowScreenTipsOnObjects.md)
- [ShowTipPages](Publisher.Options.ShowTipPages.md)
- [TypeNReplace](Publisher.Options.TypeNReplace.md)
- [UseCatalogAtStartup](Publisher.Options.UseCatalogAtStartup.md)
- [UseWizardForBlankPublication](Publisher.Options.UseWizardForBlankPublication.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]