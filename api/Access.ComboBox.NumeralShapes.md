---
title: ComboBox.NumeralShapes property (Access)
keywords: vbaac10.chm11467
f1_keywords:
- vbaac10.chm11467
ms.prod: access
api_name:
- Access.ComboBox.NumeralShapes
ms.assetid: 93cb42d2-6274-3af4-0801-87ecf8eb4252
ms.date: 03/02/2019
localization_priority: Normal
---


# ComboBox.NumeralShapes property (Access)


## Syntax

_expression_.**NumeralShapes**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **NumeralShapes** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|System|0|Numeral shapes are determined by the **Numeral Shapes** system setting.|
|Arabic|1|Arabic digit shapes are used to display and print numerals.|
|National|2|National digit shapes are used to display and print numerals.|
|Context|3|Numeral shapes are determined by Unicode context rules for adjacent text.|

## Example

The following example changes the **NumeralShapes** property for the selected control to 0 (numeral shapes will be determined by the **Numeral Shapes** system setting).

```vb
Public Sub ChangeNumeralShapes(ctl As Control) 
 ctl.NumeralShapes = 0 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]