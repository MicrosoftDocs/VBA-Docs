---
title: MatchFound property
keywords: fm20.chm5225061
f1_keywords:
- fm20.chm5225061
ms.prod: office
api_name:
- Office.MatchFound
ms.assetid: db350684-1758-a849-c9e1-34714a00f1c3
ms.date: 11/16/2018
localization_priority: Normal
---


# MatchFound property

Indicates whether the text that a user has typed into a combo box matches any of the entries in the list.

## Syntax

_object_.**MatchFound**

The **MatchFound** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Return values

The **MatchFound** property return values are:

|Value|Description|
|:-----|:-----|
|**True**|The contents of the **Value** property matches one of the records in the list.|
|**False**|The contents of **Value** does not match any of the records in the list (default).|

## Remarks

The **MatchFound** property is read-only. It is not applicable when the **MatchEntry** property is set to **fmMatchEntryNone**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]