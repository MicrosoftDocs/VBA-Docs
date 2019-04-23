---
title: ScrollBars, KeepScrollBarsVisible properties example
keywords: fm20.chm5225137
f1_keywords:
- fm20.chm5225137
ms.prod: office
ms.assetid: a935d8ab-2060-2794-69a8-ba7c8ceed3d1
ms.date: 11/14/2018
localization_priority: Normal
---


# ScrollBars, KeepScrollBarsVisible properties example

The following example uses the **[ScrollBars](scrollbars-property.md)** and the **[KeepScrollBarsVisible](keepscrollbarsvisible-property.md)** properties to add scroll bars to a page of a **[MultiPage](multipage-control.md)** and to a **[Frame](frame-control.md)**. The user chooses an option button that, in turn, specifies a value for **KeepScrollBarsVisible**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **MultiPage** named MultiPage1.    
- A **Frame** named Frame1.    
- Four **[OptionButton](optionbutton-control.md)** controls named OptionButton1 through OptionButton4.
    

```vb
Private Sub UserForm_Initialize() 
 MultiPage1.Pages(0).ScrollBars = fmScrollBarsBoth 
 MultiPage1.Pages(0).KeepScrollBarsVisible = fmScrollBarsNone 
 
 Frame1.ScrollBars = fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = fmScrollBarsNone 
 
 OptionButton1.Caption = "No scroll bars" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Horizontal scroll bars" 
 OptionButton3.Caption = "Vertical scroll bars" 
 OptionButton4.Caption = "Both scroll bars" 
End Sub 
 
Private Sub OptionButton1_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsNone 
 Frame1.KeepScrollBarsVisible = fmScrollBarsNone 
End Sub 
 
Private Sub OptionButton2_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsHorizontal 
 Frame1.KeepScrollBarsVisible = _ 
 fmScrollBarsHorizontal 
End Sub 
 
Private Sub OptionButton3_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsVertical 
 Frame1.KeepScrollBarsVisible = _ 
 fmScrollBarsVertical 
End Sub 
 
Private Sub OptionButton4_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = fmScrollBarsBoth 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]