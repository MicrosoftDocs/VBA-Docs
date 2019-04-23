---
title: MatchEntry property
keywords: fm20.chm5225060
f1_keywords:
- fm20.chm5225060
ms.prod: office
api_name:
- Office.MatchEntry
ms.assetid: 8f3ab1b9-5d69-b955-423b-be259a94a2f4
ms.date: 11/16/2018
localization_priority: Normal
---


# MatchEntry property

Returns or sets a value indicating how a **[ListBox](listbox-control.md)** or **[ComboBox](combobox-control.md)** searches its list as the user types.

## Syntax

_object_.**MatchEntry** [= _fmMatchEntry_ ]

The **MatchEntry** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmMatchEntry_|Optional. The rule used to match entries in the list.|

## Settings

The settings for _fmMatchEntry_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmMatchEntryFirstLetter_|0|Basic matching. The control searches for the next entry that starts with the character entered. Repeatedly typing the same letter [cycles](../../Glossary/glossary-vba.md#cycle) through all entries beginning with that letter.|
| _FmMatchEntryComplete_|1|Extended matching. As each character is typed, the control searches for an entry matching all characters entered (default).|
| _FmMatchEntryNone_|2|No matching.|

## Remarks

The **MatchEntry** property searches entries from the **TextColumn** property of a **ListBox** or **ComboBox**.

The control searches the column identified by **TextColumn** for an entry that matches the user's typed entry. Upon finding a match, the row containing the match is selected, the contents of the column are displayed, and the contents of its **BoundColumn** property become the value of the control. If the match is unambiguous, finding the match initiates the [Click](click-event.md) event.

The control initiates the Click event as soon as the user types a sequence of characters that match exactly one entry in the list. As the user types, the entry is compared with the current row in the list and with the next row in the list. When the entry matches only the current row, the match is unambiguous.

In Microsoft Forms, this is true regardless of whether the list is sorted. This means that the control finds the first occurrence that matches the entry, based on the order of items in the list.

Note that in either case, the matched entry is not unique; however, it is sufficiently different from the adjacent entry that the control interprets the match as unambiguous and initiates the Click event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
