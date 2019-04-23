---
title: Style property, Label control example
keywords: fm20.chm5225124
f1_keywords:
- fm20.chm5225124
ms.prod: office
ms.assetid: d2eca73d-942f-f1d0-ce04-2cbbcd36d882
ms.date: 11/14/2018
localization_priority: Normal
---


# Style property, Label control example

The following example uses the **[Style](style-property.md)** property to specify the appearance of the tabs in **[MultiPage](multipage-control.md)** and **[TabStrip](tabstrip-control.md)**. This example also demonstrates using a **[Label](label-control.md)**. The user chooses a style by selecting an **[OptionButton](optionbutton-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **Label** named Label1.    
- Three **OptionButton** controls named OptionButton1 through OptionButton3.    
- A **MultiPage** named MultiPage1.    
- A **TabStrip** named TabStrip1.   
- Any control inside the **TabStrip**.   
- Any control in each page of the **MultiPage**.
    

```vb
Private Sub OptionButton1_Click() 
 MultiPage1.Style = fmTabStyleTabs 
 TabStrip1.Style = fmTabStyleTabs 
End Sub 
 
Private Sub OptionButton2_Click() 
 'Note that the page borders are invisible 
 MultiPage1.Style = fmTabStyleButtons 
 TabStrip1.Style = fmTabStyleButtons 
End Sub 
 
Private Sub OptionButton3_Click() 
 'Note that the page borders are invisible and 
 'the page body begins where the tabs normally 
 'appear. 
 MultiPage1.Style = fmTabStyleNone 
 TabStrip1.Style = fmTabStyleNone 
End Sub 
 
Private Sub UserForm_Initialize() 
 Label1.Caption = "Page/Tab Style" 
 OptionButton1.Caption = "Tabs" 
 OptionButton1.Value = True 
 MultiPage1.Style = fmTabStyleTabs 
 TabStrip1.Style = fmTabStyleTabs 
 
 OptionButton2.Caption = "Buttons" 
 OptionButton3.Caption = "No Tabs or Buttons" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]