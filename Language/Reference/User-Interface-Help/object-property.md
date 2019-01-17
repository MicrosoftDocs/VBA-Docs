---
title: Object property
keywords: fm20.chm2001610
f1_keywords:
- fm20.chm2001610
ms.prod: office
api_name:
- Office.Object
ms.assetid: 94762c71-9ab8-98dd-5357-8ddb8b7b0156
ms.date: 11/16/2018
localization_priority: Normal
---


# Object property

Overrides a standard property or method when a new control has a property or method of the same name.

## Syntax

_object_.**Object** [. _property_ |. _method_ ]

The **Object** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. The name of an object that you have added to the Microsoft Forms Toolbox.|
| _property_|Optional. A property that has the same name as a standard Microsoft Forms property.|
| _method_|Optional. A method that has the same name as a standard Microsoft Forms method.|

## Remarks

**Object** is read-only.

If you add a new control to the Microsoft Forms Toolbox, it is possible that the added control will have a property or method with the same name as a standard Microsoft Forms property or method. The **Object** property lets you use the property or method from the added control, rather than the standard property or method.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]