---
title: CreateEventProc method (VBA Add-In Object Model)
keywords: vbob6.chm104021
f1_keywords:
- vbob6.chm104021
ms.prod: office
ms.assetid: afcdc0a2-aa3d-6882-f89c-17f0dcf3df2b
ms.date: 12/06/2018
localization_priority: Normal
---


# CreateEventProc method (VBA Add-In Object Model)

Creates an event [procedure](../../Glossary/vbe-glossary.md#procedure).

## Syntax

_object_.**CreateEventProc** (_eventname_, _objectname_) **As Long**

<br/>

The **CreateEventProc** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _eventname_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression) specifying the name of the event that you want to add to the [module](../../Glossary/vbe-glossary.md#module).|
| _objectname_|Required. A string expression specifying the name of the object that is the source of the event.|

## Remarks

Use the **CreateEventProc** method to create an event procedure. For example, to create an event procedure for the **Click** event of a **Command Button** control named `Command1`, you would use the following code, where `CM` represents an object of type **CodeModule**:

```vb
TextLocation = CM.CreateEventProc("Click", "Command1")
```

The **CreateEventProc** method returns the line at which the body of the event procedure starts. 

**CreateEventProc** fails if the [arguments](../../Glossary/vbe-glossary.md#argument) refer to a nonexistent event.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]