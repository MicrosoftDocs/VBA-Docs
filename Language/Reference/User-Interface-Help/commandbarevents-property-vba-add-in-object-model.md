---
title: CommandBarEvents property (VBA Add-In Object Model)
keywords: vbob6.chm100197
f1_keywords:
- vbob6.chm100197
ms.prod: office
ms.assetid: 342f5e9c-c5cc-4596-0b05-0985df1aba49
ms.date: 12/06/2018
---


# CommandBarEvents property (VBA Add-In Object Model)

Returns the **CommandBarEvents** object. Read-only.

## Settings

The setting for the [argument](../../Glossary/vbe-glossary.md#argument) you pass to the **CommandBarEvents** property is:

|**Argument**|**Description**|
|:-----|:-----|
| _vbcontrol_|Must be an object of type **CommandBarControl**.|

## Remarks

Use the **CommandBarEvents** property to return an [event source object](../../Glossary/vbe-glossary.md#event-source-object) that triggers an event when a command bar button is clicked. 

The argument passed to the **CommandBarEvents** property is the command bar control for which the Click event will be triggered.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)