---
title: RemoveControl event
keywords: fm20.chm2000200
f1_keywords:
- fm20.chm2000200
ms.prod: office
api_name:
- Office.RemoveControl
ms.assetid: 6e6abe85-4c0c-8fc9-668c-009e6f1a3d76
ms.date: 11/15/2018
localization_priority: Normal
---


# RemoveControl event

Occurs when a control is deleted from the [container](../../Glossary/vbe-glossary.md#container).

## Syntax

For MultiPage <br/>
**Private Sub**_object_ _**RemoveControl(**_index_**As Long**, _ctrl_**As Control)**

For all other controls <br/>
**Private Sub**_object_ _**RemoveControl(**_ctrl_**As Control)**

The **RemoveControl** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the page in a **[MultiPage](multipage-control.md)** that contained the deleted control.|
| _ctrl_|Required. The deleted control.|

## Remarks

This event occurs when a control is deleted from the form, not when a control is unloaded due to a form being closed.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]