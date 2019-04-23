---
title: Calendar property (Visual Basic for Applications)
keywords: vblr6.chm1117202
f1_keywords:
- vblr6.chm1117202
ms.prod: office
api_name:
- Office.Calendar
ms.assetid: ca321712-934e-2aee-46b8-b2895be362ea
ms.date: 12/19/2018
localization_priority: Normal
---


# Calendar property

Returns or sets a value specifying the type of calendar to use with your [project](../../Glossary/vbe-glossary.md#project).

You can use one of two settings for **Calendar**.

|Setting|Value|Description|
|:-----|:-----|:-----|
|**vbCalGreg**|0|Use Gregorian calendar (default).|
|**vbCalHijri**|1|Use Hijri calendar.|

## Remarks

You can only set the **Calendar** property programmatically. For example, to use the Hijri calendar, use:

```vb
Calendar = vbCalHijri

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]