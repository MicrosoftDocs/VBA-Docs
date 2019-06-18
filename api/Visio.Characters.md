---
title: Characters object (Visio)
keywords: vis_sdr.chm10050
f1_keywords:
- vis_sdr.chm10050
ms.prod: visio
api_name:
- Visio.Characters
ms.assetid: aaff009b-c665-c2ea-8494-e917126d8491
ms.date: 06/19/2019
localization_priority: Normal
---


# Characters object (Visio)

Represents a shape's text with the text fields expanded to the number of characters they display in a drawing window.


## Remarks

To retrieve a **Characters** object, use the **[Characters](visio.shape.characters.md)** property of a **Shape** object.

The default property of a **Characters** object is **Text**.

The **Begin** and **End** properties of a **Characters** object determine the range of the shape's text that is represented by the **Characters** object. Initially, the range contains all of the shape's text; you can set the **Begin** and **End** properties to specify a subrange of the text.

After you retrieve a **Characters** object, you can use its **Text** property to retrieve or set the shape's text. Use the **Copy**, **Cut**, or **Paste** method to copy, cut, or paste the **Characters** object's text to or from the Clipboard. 

Use the **CharProps** or **ParaProps** property to change the **Characters** object's formatting.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this object maps to the following types:

- **Microsoft.Office.Interop.Visio.IVCharacters**
    

## Events

-  [TextChanged](Visio.Characters.TextChanged.md)

## Methods

-  [AddCustomField](Visio.Characters.AddCustomField.md)
-  [AddCustomFieldU](Visio.Characters.AddCustomFieldU.md)
-  [AddField](Visio.Characters.AddField.md)
-  [AddFieldEx](Visio.Characters.AddFieldEx.md)
-  [Copy](Visio.Characters.Copy.md)
-  [Cut](Visio.Characters.Cut.md)
-  [Delete](Visio.Characters.Delete.md)
-  [Paste](Visio.Characters.Paste.md)

## Properties

-  [Application](Visio.Characters.Application.md)
-  [Begin](Visio.Characters.Begin.md)
-  [CharCount](Visio.Characters.CharCount.md)
-  [CharProps](Visio.Characters.CharProps.md)
-  [CharPropsRow](Visio.Characters.CharPropsRow.md)
-  [ContainingMasterID](Visio.Characters.ContainingMasterID.md)
-  [ContainingPageID](Visio.Characters.ContainingPageID.md)
-  [Document](Visio.Characters.Document.md)
-  [End](Visio.Characters.End.md)
-  [EventList](Visio.Characters.EventList.md)
-  [FieldCategory](Visio.Characters.FieldCategory.md)
-  [FieldCode](Visio.Characters.FieldCode.md)
-  [FieldFormat](Visio.Characters.FieldFormat.md)
-  [FieldFormula](Visio.Characters.FieldFormula.md)
-  [FieldFormulaU](Visio.Characters.FieldFormulaU.md)
-  [IsField](Visio.Characters.IsField.md)
-  [ObjectType](Visio.Characters.ObjectType.md)
-  [ParaProps](Visio.Characters.ParaProps.md)
-  [ParaPropsRow](Visio.Characters.ParaPropsRow.md)
-  [PersistsEvents](Visio.Characters.PersistsEvents.md)
-  [RunBegin](Visio.Characters.RunBegin.md)
-  [RunEnd](Visio.Characters.RunEnd.md)
-  [Shape](Visio.Characters.Shape.md)
-  [Stat](Visio.Characters.Stat.md)
-  [TabPropsRow](Visio.Characters.TabPropsRow.md)
-  [Text](Visio.Characters.Text.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]