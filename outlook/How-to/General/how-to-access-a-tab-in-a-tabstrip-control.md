---
title: "How to: Access a Tab in a TabStrip Control"
keywords: olfm10.chm3077151
f1_keywords:
- olfm10.chm3077151
ms.prod: outlook
ms.assetid: 29aba68e-7123-2c41-795f-7bdba8d1b89f
ms.date: 06/08/2017
localization_priority: Normal
---


# Access a Tab in a TabStrip Control

The following example accesses an individual tab of a  **[TabStrip](../../../api/Outlook.tabstrip.md)** in several ways:


- Using the  **[Tabs](../../../api/Outlook.tabs.md)** collection with a numeric index.
    
- Using the name of the individual  **[Tab](../../../api/Outlook.tab.md)**.
    
- Using the  **[SelectedItem](../../../api/Outlook.tabstrip.selecteditem.md)** property.
    

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event of the item will activate. Make sure that the form contains a **TabStrip** named TabStrip1.




```vb
Sub Item_Open() 
 Dim TabStrip1 
 Dim TabName 
 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TabStrip1") 
 For i = 0 To TabStrip1.Count - 1 
 'Using index (numeric or string) 
 MsgBox "TabStrip1.Tabs(i).Caption = " & TabStrip1.Tabs(i).Caption 
 MsgBox "TabStrip1.Tabs.Item(i).Caption = " & TabStrip1.Tabs.Item(i).Caption 
 
 'Use Tab object without referring to Tabs collection 
 If i = 0 Then 
 MsgBox "TabStrip1.Tab1. Caption = " & TabStrip1.Tab1.Caption 
 ElseIf i = 1 Then 
 MsgBox "TabStrip1.Tab2. Caption = " & TabStrip1.Tab2.Caption 
 End If 
 
 'Use SelectedItem Property 
 TabStrip1.Value = i 
 MsgBox " TabStrip1.SelectedItem.Caption = " & TabStrip1.SelectedItem.Caption 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]