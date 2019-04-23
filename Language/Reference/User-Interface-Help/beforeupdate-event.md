---
title: BeforeUpdate event
keywords: fm20.chm2000050
f1_keywords:
- fm20.chm2000050
ms.prod: office
api_name:
- Office.BeforeUpdate
ms.assetid: ccf0fa5d-a069-cba6-5725-072b141fa80b
ms.date: 11/15/2018
localization_priority: Normal
---


# BeforeUpdate event

Occurs before data in a control is changed.

## Syntax

**Private Sub**_object_ _**BeforeUpdate( ByVal**_Cancel_**As MSForms.ReturnBoolean)**

The **BeforeUpdate** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Cancel_|Required. Event status. **False** indicates that the control should handle the event (default). **True** cancels the update and indicates the application should handle the event.|

## Remarks

The BeforeUpdate event occurs regardless of whether the control is [bound](../../Glossary/glossary-vba.md#bound) (that is, when the **RowSource** property specifies a [data source](../../Glossary/glossary-vba.md#data-source) for the control). 

This event occurs before the AfterUpdate and Exit events for the control (and before the Enter event for the next control that receives [focus](../../Glossary/vbe-glossary.md#focus)).

If you set the _Cancel_ argument to **True**, the focus remains on the control and neither the AfterUpdate event nor the Exit event occurs.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]