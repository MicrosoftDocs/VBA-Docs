---
title: Using For Each...Next statements (VBA)
keywords: vbcn6.chm1076683
f1_keywords:
- vbcn6.chm1076683
ms.prod: office
ms.assetid: 76df8944-219a-c28b-c449-39a3108c11be
ms.date: 12/26/2018
localization_priority: Normal
---


# Using For Each...Next statements

**[For Each...Next](../../reference/user-interface-help/for-eachnext-statement.md)** statements repeat a block of [statements](../../Glossary/vbe-glossary.md#statement) for each [object](../../Glossary/vbe-glossary.md#object) in a [collection](../../Glossary/vbe-glossary.md#collection) or each element in an [array](../../Glossary/vbe-glossary.md#array). Visual Basic automatically sets a [variable](../../Glossary/vbe-glossary.md#variable) each time the loop runs. For example, the following [procedure](../../Glossary/vbe-glossary.md#procedure) closes all forms except the form containing the procedure that's running.

```vb
Sub CloseForms() 
 For Each frm In Application.Forms 
 If frm.Caption <> Screen. ActiveForm.Caption Then frm.Close 
 Next 
End Sub
```

The following code loops through each element in an array and sets the value of each to the value of the index variable I.

```vb
Dim TestArray(10) As Integer, I As Variant 
For Each I In TestArray 
 TestArray(I) = I 
Next I 

```


## Looping through a range of cells

Use a **For Each...Next** loop to loop through the cells in a range. The following procedure loops through the range A1:D10 on Sheet1 and sets any number whose absolute value is less than 0.01 to 0 (zero).

```vb
Sub RoundToZero() 
 For Each myObject in myCollection 
 If Abs(myObject.Value) < 0.01 Then myObject.Value = 0 
 Next 
End Sub
```

## Exiting a For Each...Next loop before it is finished

You can exit a **For Each...Next** loop by using the **[Exit For](../../reference/user-interface-help/exit-statement.md)** statement. For example, when an error occurs, use the **Exit For** statement in the **True** statement block of either an **[If...Then...Else](../../reference/user-interface-help/ifthenelse-statement.md)** statement or a **[Select Case](../../reference/user-interface-help/select-case-statement.md)** statement that specifically checks for the error. If the error does not occur, the **If…Then…Else** statement is **False** and the loop continues to run as expected.

The following example tests for the first cell in the range A1:B5 that does not contain a number. If such a cell is found, a message is displayed and **Exit For** exits the loop.

```vb
Sub TestForNumbers() 
 For Each myObject In MyCollection 
 If IsNumeric(myObject.Value) = False Then 
 MsgBox "Object contains a non-numeric value." 
 Exit For 
 End If 
 Next c 
End Sub
```


## Using For Each...Next loop to iterate over a [VBA class]() you have written

**For Each...Next** loops don't only iterate over arrays, & instances of the [**Collection** object/class](../../reference/user-interface-help/collection-object.md). **For Each...Next** loops can iterate over a VBA class you have written yourself.

Below is an example demonstrating how you can do this.

1) First of all, create a [class module](../../glossary/vbe-glossary.md#class-module) in the VBE (Visual Basic Editor), and rename it 'CustomCollection'<sup> [cc1](#cc1)</sup>.

2) Place the following code in the newly created module:

```vb
Private MyCollection As New Collection

' The Initialize event automatically gets triggered
' when instances of this class are created.
' It then triggers the execution of this procedure.
Private Sub Class_Initialize()
    With MyCollection
        .Add "First Item"
        .Add "Second Item"
        .Add "Third Item"
    End With
End Sub

' Property Get procedure for the setting up of
' this class so that it works with 'For Each...'
' constructs.
Property Get NewEnum() As IUnknown
' Attribute NewEnum.VB_UserMemId = -4

   Set NewEnum = MyCollection.[_NewEnum]
End Property
```

3) Next, export this module to a file and store it locally<sup> [cc2](#cc2)</sup>.

4) Once exported, open the exported file using a text editor (Window's _Notepad_ software should be sufficient.) The file contents should look like the following:

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyCollection As New Collection

' The Initialize event automatically gets triggered
' when instances of this class are created.
' It then triggers the execution of this procedure.
Private Sub Class_Initialize()
    With MyCollection
        .Add "First Item"
        .Add "Second Item"
        .Add "Third Item"
    End With
End Sub

' Property Get procedure for the setting up of
' this class so that it works with 'For Each...'
' constructs.
Property Get NewEnum() As IUnknown
' Attribute NewEnum.VB_UserMemId = -4

   Set NewEnum = MyCollection.[_NewEnum]
End Property
```

5) Now using the text editor, remove the `'` character from the ***first line under*** the "`Property Get NewEnum() As IUnknown`" text in the file. Save the modified file.

6) Back in the VBE, remove the class you created, from your VBA project & don't choose to export it when prompted<sup> [cc3](#cc3)</sup>.

7) Next, import the file that you removed the `'` character from, back into the VBE<sup> [cc4](#cc4)</sup>.

8) Finally, run the following code to see that you can now iterate over your custom VBA class that you have written using both the VBE & a text editor:

```vbe
 Dim Element
 Dim MyCustomCollection As New CustomCollection
 For Each Element In MyCustomCollection
  MsgBox Element
 Next
```
<table>
<thead>
<tr>
<td colspan=2>
<sup><strong>Footnotes</strong></sup>
</td>
</tr>
</thead>
<tr>
<td valign="top" align="right">
<sup><a name="cc1"><strong>[cc1]</strong></a></sup>
</td>
<td>
<sup>You can create a <a href="../../glossary/vbe-glossary.md#class-module">class module</a> by choosing the <em>Class Module</em> menu item from the <a href="../../reference/user-interface-help/insert-menu.md"><em>Insert</em> menu</a>. You can rename a class module by modifying its properties from the <a href="../../reference/user-interface-help/use-the-properties-window.md">properties window</a>.</sup>
</td>
</tr>
<tr>
<td valign="top" align="right">
<sup><a name="cc2"><strong>[cc2]</strong></a></sup>
</td>
<td>
<sup>You can activate the <a href="../../reference/user-interface-help/export-file-dialog-box.md">Export File dialog box</a> by choosing the <a href="../../reference/user-interface-help/file-menu.md#import-file-export-file"><em>Export File...</em> menu item</a> from the <a href="../../reference/user-interface-help/file-menu.md"><em>File_ menu</em>.</sup>
</td>
</tr><tr>
<td valign="top" align="right">
<sup><a name="cc3"><strong>[cc3]</strong></a></sup>
</td>
<td>
<sup>You can remove a class module from the VBE by choosing the <a href="../../reference/user-interface-help/file-menu.md#remove-item"><em>Remove</em> Item menu item</a> from the <a href="../../reference/user-interface-help/file-menu.md"><em>File</em> menu</a>.</sup>
</td>
</tr><tr>
<td valign="top" align="right">
<sup><a name="cc4"><strong>[cc4]</strong></a></sup>
</td>
<td>
<sup>You can import an external class-module file, by activating the <a href="../../reference/user-interface-help/import-file-dialog-box.md">Import File dialog box</a> by choosing the <a href="../../reference/user-interface-help/file-menu.md#import-file-export-file"><em>Import File...</em> menu item</a> from the <a href="../../reference/user-interface-help/file-menu.md"><em>File</em> menu</a>.</sup>
</td>
</tr>
</table>

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
