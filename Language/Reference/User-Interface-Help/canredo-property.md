---
title: CanRedo property
keywords: fm20.chm2000860
f1_keywords:
- fm20.chm2000860
ms.prod: office
api_name:
- Office.CanRedo
ms.assetid: 18b4b51d-3a8a-e03d-14b2-b262f6a12c78
ms.date: 11/15/2018
localization_priority: Normal
---


# CanRedo property

Indicates whether the most recent Undo can be reversed.

## Syntax

_object_.**CanRedo**

The **CanRedo** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Return values

The **CanRedo** property return values are:

|Value|Description|
|:-----|:-----|
|**True**|The most recent Undo can be reversed.|
|**False**|The most recent Undo is irreversible.|

## Remarks

**CanRedo** is read-only.

To Redo an action means to reverse an Undo; it does not necessarily mean to repeat the last user action.

The following user actions illustrate using Undo and Redo:

- Change the setting of an option button.   
- Enter text into a text box.    
- Click **Undo**. The text disappears from the text box.   
- Click **Undo**. The option button reverts to its previous setting.   
- Click **Redo**. The value of the option button changes.   
- Click **Redo**. The text reappears in the text box.
    
## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]