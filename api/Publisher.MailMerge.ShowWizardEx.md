---
title: MailMerge.ShowWizardEx method (Publisher)
keywords: vbapb10.chm6225944
f1_keywords:
- vbapb10.chm6225944
ms.prod: publisher
api_name:
- Publisher.MailMerge.ShowWizardEx
ms.assetid: 3815204f-5f09-5a25-a2e4-5de4889c9919
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.ShowWizardEx method (Publisher)

Displays the specified catalog or mail merge wizard in a document.


## Syntax

_expression_.**ShowWizardEx** (_ShowDocumentStep_, _ShowTemplateStep_, _ShowDataStep_, _ShowWriteStep_, _ShowPreviewStep_, _ShowMergeStep_, _MergeType_, _iStep_)

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ShowDocumentStep_|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions, **True** (the default) displayed the "Select a merge type" step; **False** removed the step.|
|_ShowTemplateStep_|Optional| **Boolean**| This parameter does not apply to Microsoft Publisher.|
|_ShowDataStep_|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions, **True** (the default) displayed the "Select data source" step; **False** removed the step.|
|_ShowWriteStep_|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions, **True** (the default) displayed the "Create your publication" step; **False** removed the step.|
|_ShowPreviewStep_|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions, **True** (the default) displayed the "Preview your publication" step; **False** removed the step.|
|_ShowMergeStep_|Optional| **Boolean**|Not used in Microsoft Publisher 2007. In previous versions, **True** (the default) displayed the "Complete the merge" step; **False** removed the step.|
|_MergeType_|Optional| **[PbMergeType](Publisher.PbMergeType.md)** |The merge type to use. Can be one of the **PbMergeType** constants declared in the Microsoft Publisher type library. The default is **pbMergeDefault**.|
|_iStep_|Optional| **Long**|The initial step. See Remarks for information about default values.|

## Remarks

Passing **pbMergeDefault** for _MergeType_ starts a new mail merge; if the publication is already a merge, it leaves the merge type unchanged.

Passing a merge type that is different from the current publication's merge type changes the publication to that new type of merge, but disconnects the data source. Doing so results in the loss of previously inserted fields when the change is to or from a catalog merge type.

Wizard steps correspond to the sequence of merge task panes in the user interface. If no data source is connected, the merge wizard always starts on the first step (the first task pane). If a data source is connected, the wizard starts on Step 2 by default, unless you use the _iStep_ parameter to specify starting with Step 1 or Step 3.


## Example

This example checks whether the mail merge wizard is closed, and if it is, displays it.

```vb
Public Sub ShowWizardEx_Example() 
 With ActiveDocument.MailMerge 
 If .WizardState = 0 Then 
 .ShowWizardEx 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]