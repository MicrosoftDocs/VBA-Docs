---
title: MultiLine property 
keywords: fm20.chm2001560
f1_keywords:
- fm20.chm2001560
ms.prod: office
api_name:
- Office.MultiLine
ms.assetid: eadbbea9-f4ab-bb60-dff8-950d03b70842
ms.date: 11/16/2018
localization_priority: Normal
---


# MultiLine property 

Specifies whether a control can accept and display multiple lines of text.

## Syntax

_object_.**MultiLine** [= _Boolean_ ]

The **MultiLine** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control supports more than one line of text.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|The text is displayed across multiple lines (default).|
|**False**|The text is not displayed across multiple lines.|

## Remarks

A multiline **[TextBox](textbox-control.md)** allows absolute line breaks and adjusts its quantity of lines to accommodate the amount of text it holds. If needed, a multiline control can have vertical scroll bars.

A single-line **TextBox** doesn't allow absolute line breaks and doesn't use vertical scroll bars.
Single-line controls ignore the value of the **WordWrap** property.

> [!NOTE] 
> If you change **MultiLine** to **False** in a multiline **TextBox**, all the characters in the **TextBox** will be combined into one line, including non-printing characters (such as carriage returns and new-lines).


## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]