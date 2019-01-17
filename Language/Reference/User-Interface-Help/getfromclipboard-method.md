---
title: GetFromClipboard method (Microsoft Forms)
keywords: fm20.chm5224960
f1_keywords:
- fm20.chm5224960
ms.prod: office
api_name:
- Office.GetFromClipboard
ms.assetid: 8a034bf7-b6cf-ed9f-2e1c-2a4325494970
ms.date: 11/15/2018
localization_priority: Normal
---


# GetFromClipboard method (Microsoft Forms)

Copies data from the Clipboard to a **[DataObject](dataobject-object.md)**.

## Syntax

_String = object_. **GetFromClipboard( )**

The **GetFromClipboard** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object name.|

## Remarks

The **DataObject** can contain multiple data items, but each item must be in a different format. For example, the **DataObject** might include one text item and one item in a custom format; but cannot include two text items.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]