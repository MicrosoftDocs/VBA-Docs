---
title: Parent property (Microsoft Forms)
keywords: fm20.chm5225075
f1_keywords:
- fm20.chm5225075
ms.prod: office
ms.assetid: a8289266-cb45-8458-ba09-c0efd19665f9
ms.date: 11/16/2018
localization_priority: Normal
---


# Parent property (Microsoft Forms)

Returns the name of the form, object, or [collection](../../Glossary/vbe-glossary.md#collection) that contains a specific control, object, or collection.

## Syntax

_object_.**Parent**

The **Parent** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|

## Remarks

**Parent** is read-only.

Use the **Parent** property to access the properties, methods, or controls of an object's parent.

This property is useful in an application in which you pass objects as arguments. For example, you could pass a control variable to a general procedure in a [module](../../Glossary/vbe-glossary.md#module), and use **Parent** to access its parent form.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]