---
title: ListBox control, List property example
keywords: fm20.chm5225171
f1_keywords:
- fm20.chm5225171
ms.prod: office
ms.assetid: 14396c81-9137-7352-906c-acf70e9e77b0
ms.date: 11/14/2018
localization_priority: Normal
---


# ListBox control, List property example

The following example swaps columns of a multicolumn **[ListBox](listbox-control.md)**. The sample uses the **[List](list-property-microsoft-forms.md)** property in two ways:

- To access and exchange individual values in the **ListBox**. In this usage, **List** has subscripts to designate the row and column of a specified value.
    
- To initially load the **ListBox** with values from an array. In this usage, **List** has no subscripts.
    
To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a **ListBox** named ListBox1 and a **[CommandButton](commandbutton-control.md)** named CommandButton1.



```vb
Dim MyArray(6, 3) 
'Array containing column values for ListBox. 
 
Private Sub UserForm_Initialize() 
 Dim i As Single 
 
 ListBox1.ColumnCount = 3 
'This list box contains 3 data columns 
 
 'Load integer values MyArray 
 For i = 0 To 5 
 MyArray(i, 0) = i 
 MyArray(i, 1) = Rnd 
 MyArray(i, 2) = Rnd 
 Next i 
 
 'Load ListBox1 
 ListBox1.List() = MyArray 
 
End Sub
```

<br/>


```vb
Private Sub CommandButton1_Click() 
' Exchange contents of columns 1 and 3 
 
 Dim i As Single 
 Dim Temp As Single 
 
 For i = 0 To 5 
 Temp = ListBox1.List(i, 0) 
 ListBox1.List(i, 0) = ListBox1.List(i, 2) 
 ListBox1.List(i, 2) = Temp 
 Next i 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
