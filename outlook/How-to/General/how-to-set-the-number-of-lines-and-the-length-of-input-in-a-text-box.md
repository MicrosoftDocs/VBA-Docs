---
title: "How to: Set the Number of Lines and the Length of Input in a Text Box"
keywords: olfm10.chm3077204
f1_keywords:
- olfm10.chm3077204
ms.prod: outlook
ms.assetid: 1b56aff7-ab6f-b595-781d-a60d0dffe7a9
ms.date: 06/08/2019
localization_priority: Normal
---


# Set the Number of Lines and the Length of Input in a Text Box

The following example counts the characters and the number of lines of text in a **TextBox](../../../api/Outlook.textbox.md)** by using the **[LineCount](../../../api/Outlook.textbox.linecount.md)** and **[TextLength](../../../api/Outlook.textbox.textlength.md)** properties, and the **SetFocus** method. In this example, the user can type into a **TextBox**, and can retrieve current values of the **neCount** and **TextLength** properties.


 **Note** The **etFocus** method is inherited from the Microsoft Forms 2.0 **TextBox** control.


To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the **pen** event will activate. Make sure that the form contains the following controls:


- A **extBox** named TextBox1.
    
- A **CommandButton](../../../api/Outlook.commandbutton.md)** named CommandButton1.
    
- Two **Label](../../../api/Outlook.label.md)** controls named Label1 and Label2.
    



```vb
'Type SHIFT+ENTER to start a new line in the text box. 
 
Dim CommandButton1 
Dim TextBox1 
Dim Label1 
Dim Label2 
 
Sub CommandButton1_Click() 
 'Must first give TextBox1 the focus to get line count 
 TextBox1.SetFocus 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 Set Label2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("Label2") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CommandButton1") 
 
 CommandButton1.WordWrap = True 
 CommandButton1.AutoSize = True 
 CommandButton1.Caption = "Get Counts" 
 
 Label1.Caption = "LineCount = " 
 Label2.Caption = "TextLength = " 
 
 TextBox1.MultiLine = True 
 TextBox1.WordWrap = True 
 TextBox1.Text = "Enter your text here." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]