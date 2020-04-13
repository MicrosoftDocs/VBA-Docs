---
title: OLEFormat object (Word)
keywords: vbawd10.chm2355
f1_keywords:
- vbawd10.chm2355
ms.prod: word
api_name:
- Word.OLEFormat
ms.assetid: d4c7aa65-5d3a-0b79-914b-6f908b506f63
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEFormat object (Word)

Represents the OLE characteristics (other than linking) for an OLE object, ActiveX control, or field.


## Remarks

Use the **OLEFormat** property for a shape, inline shape, or field to return the **OLEFormat** object. The following example displays the class type for the first shape on the active document.


```vb
MsgBox ActiveDocument.Shapes(1).OLEFormat.ClassType
```

Not all types of shapes, inline shapes, and fields have OLE capabilities. Use the **Type** property for the **Shape** and **InlineShape** objects to determine what category the specified shape or inline shape falls into. The **Type** property for a **Field** object returns the type of field.

You can use the **Activate**, **Edit**, **Open**, and **DoVerb** methods to automate an OLE object.

Use the **Object** property to return an object that represents an ActiveX control or OLE object. With this object, you can use the properties and methods of the container application or the ActiveX control.


## Methods



|Name|
|:-----|
|[Activate](Word.OLEFormat.Activate.md)|
|[ActivateAs](Word.OLEFormat.ActivateAs.md)|
|[ConvertTo](Word.OLEFormat.ConvertTo.md)|
|[DoVerb](Word.OLEFormat.DoVerb.md)|
|[Edit](Word.OLEFormat.Edit.md)|
|[Open](Word.OLEFormat.Open.md)|

## Properties



|Name|
|:-----|
|[Application](Word.OLEFormat.Application.md)|
|[ClassType](Word.OLEFormat.ClassType.md)|
|[Creator](Word.OLEFormat.Creator.md)|
|[DisplayAsIcon](Word.OLEFormat.DisplayAsIcon.md)|
|[IconIndex](Word.OLEFormat.IconIndex.md)|
|[IconLabel](Word.OLEFormat.IconLabel.md)|
|[IconName](Word.OLEFormat.IconName.md)|
|[IconPath](Word.OLEFormat.IconPath.md)|
|[Label](Word.OLEFormat.Label.md)|
|[Object](Word.OLEFormat.Object.md)|
|[Parent](Word.OLEFormat.Parent.md)|
|[PreserveFormattingOnUpdate](Word.OLEFormat.PreserveFormattingOnUpdate.md)|
|[ProgID](Word.OLEFormat.ProgID.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]