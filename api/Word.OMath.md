---
title: OMath object (Word)
keywords: vbawd10.chm2691
f1_keywords:
- vbawd10.chm2691
ms.prod: word
api_name:
- Word.OMath
ms.assetid: 82f2f81b-e2d5-140f-bdcc-8b52b821b24d
ms.date: 06/08/2017
localization_priority: Normal
---


# OMath object (Word)

Represents an equation.  **OMath** objects are members of the **OMaths** collection.


## Remarks

Use the **Add** method of the **OMaths** collection to create an equation and add it to a document, selection, or range. The following example creates an equation and uses the **BuildUp** method to convert the equation to professional format.


```vb
Dim objRange As Range 
Dim objEq As OMath 
 
Set objRange = Selection.Range 
objRange.Text = "Celsius = (5/9)(Fahrenheit - 32)" 
Set objRange = Selection.OMaths.Add(objRange) 
Set objEq = objRange.OMaths(1) 
objEq.BuildUp
```


## Methods



|Name|
|:-----|
|[BuildUp](Word.OMath.BuildUp.md)|
|[ConvertToLiteralText](Word.OMath.ConvertToLiteralText.md)|
|[ConvertToMathText](Word.OMath.ConvertToMathText.md)|
|[ConvertToNormalText](Word.OMath.ConvertToNormalText.md)|
|[Linearize](Word.OMath.Linearize.md)|
|[Remove](Word.OMath.Remove.md)|

## Properties



|Name|
|:-----|
|[AlignPoint](Word.OMath.AlignPoint.md)|
|[Application](Word.OMath.Application.md)|
|[ArgIndex](Word.OMath.ArgIndex.md)|
|[ArgSize](Word.OMath.ArgSize.md)|
|[Breaks](Word.OMath.Breaks.md)|
|[Creator](Word.OMath.Creator.md)|
|[Functions](Word.OMath.Functions.md)|
|[Justification](Word.OMath.Justification.md)|
|[NestingLevel](Word.OMath.NestingLevel.md)|
|[Parent](Word.OMath.Parent.md)|
|[ParentArg](Word.OMath.ParentArg.md)|
|[ParentCol](Word.OMath.ParentCol.md)|
|[ParentFunction](Word.OMath.ParentFunction.md)|
|[ParentOMath](Word.OMath.ParentOMath.md)|
|[ParentRow](Word.OMath.ParentRow.md)|
|[Range](Word.OMath.Range.md)|
|[Type](Word.OMath.Type.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]