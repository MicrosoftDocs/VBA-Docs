---
title: InsideHeight, InsideWidth properties
keywords: fm20.chm5225045
f1_keywords:
- fm20.chm5225045
ms.prod: office
ms.assetid: 8db4373d-0807-ec2a-f9df-37ebcbf8ef47
ms.date: 11/16/2018
localization_priority: Normal
---


# InsideHeight, InsideWidth properties

**InsideHeight** returns the height, in [points](../../Glossary/vbe-glossary.md#point), of the [client region](../../Glossary/glossary-vba.md#client-region) inside a form. **InsideWidth** returns the width, in points, of the client region inside a form.

## Syntax

_object_.**InsideHeight** <br/>
_object_.**InsideWidth**

The **InsideHeight** and **InsideWidth** property syntaxes have these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

The **InsideHeight** and **InsideWidth** properties are read-only. If the region includes a scroll bar, the returned value does not include the height or width of the scroll bar.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]