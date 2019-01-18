---
title: ColumnHeads property
keywords: fm20.chm5225021
f1_keywords:
- fm20.chm5225021
ms.prod: office
api_name:
- Office.ColumnHeads
ms.assetid: 55cd26ad-8ef3-8e65-f655-315af620658d
ms.date: 11/15/2018
localization_priority: Normal
---


# ColumnHeads property

Displays a single row of column headings for list boxes, combo boxes, and objects that accept column headings.

## Syntax

_object_.**ColumnHeads** [= _Boolean_ ]

The **ColumnHeads** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Specifies whether the column headings are displayed.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|Display column headings.|
|**False**|Do not display column headings (default).|

Headings in combo boxes appear only when the list drops down.

## Remarks

When the system uses the first row of data items as column headings, they can't be selected.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]