---
title: OlkControl object (Outlook)
keywords: vbaol11.chm1000510
f1_keywords:
- vbaol11.chm1000510
ms.prod: outlook
ms.assetid: 426a3ce8-9103-d72e-13ee-9fb47ae0eb07
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkControl object (Outlook)

Defines a set of control properties common to some Microsoft Outlook controls.


## Remarks

The members offered by  **OlkControl** can apply to most Outlook controls. **OlkControl** provides a class to which you can conveniently cast an Outlook control without resorting to reflection. Although **OlkControl** does not apply to Microsoft Forms 2.0 controls, similar properties are available to Forms 2.0 controls. 

## Example

The following code sample uses the  **[OlkControl](Outlook.olkcontrol.md)** class to enable automatic resizing of a text box control with respect to any resizing of the form. It uses casting in Visual Basic to allow the text box control to use the properties of **OlkControl**.


```vb
Sub ResizeWithForm() 
 Dim myTextBox As OlkTextBox 
 Dim olkCtrl As OlkControl 
 
 ' Let the text box control use the properties of OlkControl 
 Set olkCtrl = myTextBox 
 
 ' Enable automatic adjustments of the layout with respect to the rest of the form 
 olkCtrl.EnableAutoLayout = True 
 
 ' Allow resizing the text box control horizontally and vertically with the form 
 olkCtrl.HorizontalLayout = olHorizontalLayoutGrow 
 olkCtrl.VerticalLayout = olVerticalLayoutGrow 
End Sub
```


## Properties



|Name|
|:-----|
|[ControlProperty](Outlook.OlkControl.ControlProperty.md)|
|[EnableAutoLayout](Outlook.OlkControl.EnableAutoLayout.md)|
|[Format](Outlook.OlkControl.Format.md)|
|[HorizontalLayout](Outlook.OlkControl.HorizontalLayout.md)|
|[ItemProperty](Outlook.OlkControl.ItemProperty.md)|
|[MinimumHeight](Outlook.OlkControl.MinimumHeight.md)|
|[MinimumWidth](Outlook.OlkControl.MinimumWidth.md)|
|[PossibleValues](Outlook.OlkControl.PossibleValues.md)|
|[VerticalLayout](Outlook.OlkControl.VerticalLayout.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]