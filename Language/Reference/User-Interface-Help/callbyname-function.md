---
title: CallByName function (Visual Basic for Applications)
keywords: vblr6.chm1020905
f1_keywords:
- vblr6.chm1020905
ms.prod: office
ms.assetid: e76dece5-244f-9514-4ccf-d993d6476061
ms.date: 12/11/2018
localization_priority: Normal
---


# CallByName function

Executes a method of an object, or sets or returns a property of an [object](../../Glossary/vbe-glossary.md#object).

## Syntax

**CallByName** (_object_, _procname_, _calltype_, [args()]_)

<br/>

The **CallByName** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_object_ |Required: **Variant** (**Object**). The name of the object on which the function will be executed.|
|_procname_ |Required: **Variant** (**String**). A string expression containing the name of a property or method of the object.|
|_calltype_ |Required: **Constant**. A constant of type **vbCallType** representing the type of procedure being called.|
| _args()_|Optional: **Variant** (**Array**).|

## Remarks

The **CallByName** function is used to get or set a property, or invoke a method at run time by using a string name.

In the following example, the first line uses **CallByName** to set the **[MousePointer](mousepointer-property.md)** property of a text box, the second line gets the value of the **MousePointer** property, and the third line invokes the **[Move](move-method-filesystemobject-object.md)** method to move the text box.

```vb
CallByName Text1, "MousePointer", vbLet, vbCrosshair
Result = CallByName (Text1, "MousePointer", vbGet)
CallByName Text1, "Move", vbMethod, 100, 100
```

## Example

This example uses the **CallByName** function to invoke the **Move** method of a **Command** button.

The example also uses a form (`Form1`) with a button (`Command1`), and a label (`Label1`). When the form is loaded, the **[Caption](caption-propert-microsoft-forms.md)** property of the label is set to "Move", and the name of the method to invoke. When you click the button, the **CallByName** function invokes the method to change the location of the button.

```vb
Option Explicit

Private Sub Form_Load()
    Label1.Caption = "Move"        ' Name of Move method.
End Sub

Private Sub Command1_Click()
    If Command1.Left <> 0 Then
        CallByName Command1, Label1.Caption, vbMethod, 0, 0
    Else
        CallByName Command1, Label1.Caption, vbMethod, 500, 500
    End If
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
