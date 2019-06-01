---
title: PhoneticGuide object (Publisher)
keywords: vbapb10.chm6225919
f1_keywords:
- vbapb10.chm6225919
ms.prod: publisher
api_name:
- Publisher.PhoneticGuide
ms.assetid: 164e8b54-4bad-4de9-bf6e-52c5687dfbc6
ms.date: 06/01/2019
localization_priority: Normal
---


# PhoneticGuide object (Publisher)

Represents base text with supplementary text appearing above it as a guide to pronunciation.
 
## Remarks

Use the **[PhoneticGuide](publisher.field.phoneticguide.md)** property of a **Field** object to return an existing **PhoneticGuide** object. 

Use the **[AddPhoneticGuide](publisher.fields.addphoneticguide.md)** method of a **Fields** collection to create a new **PhoneticGuide** object.

## Example

The following example adds a new **PhoneticGuide** object to the active publication.

```vb
Selection.TextRange.Fields.AddPhoneticGuide _ 
 Range:=Selection.TextRange, Text:="ver-E nIs", _ 
 Alignment:=pbPhoneticGuideAlignmentCenter, _ 
 Raise:=11, FontSize:=7
```


## Methods

- [Clear](Publisher.PhoneticGuide.Clear.md)

## Properties

- [Alignment](Publisher.PhoneticGuide.Alignment.md)
- [Application](Publisher.PhoneticGuide.Application.md)
- [BaseText](Publisher.PhoneticGuide.BaseText.md)
- [FontName](Publisher.PhoneticGuide.FontName.md)
- [FontSize](Publisher.PhoneticGuide.FontSize.md)
- [Parent](Publisher.PhoneticGuide.Parent.md)
- [Raise](Publisher.PhoneticGuide.Raise.md)
- [Text](Publisher.PhoneticGuide.Text.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]