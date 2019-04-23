---
title: Set form, report, and control properties in Visual Basic
keywords: vbaac10.chm5188061
f1_keywords:
- vbaac10.chm5188061
ms.prod: access
ms.assetid: 1f5b5f6b-b424-f35e-4add-21c45b5d74c4
ms.date: 09/26/2018
localization_priority: Normal
---


# Set form, report, and control properties in Visual Basic

**[Form](../../../api/Access.Form.md)**, **[Report](../../../api/Access.Report.md)**, and **[Control](../../../api/Access.Control.md)** objects are Microsoft Access objects. You can set properties for these objects from within a **Sub**, **Function**, or event procedure. You can also set properties for form and report sections.

## Set a property of a form or report

Refer to the individual form or report within the **[Forms](../../../api/Access.Forms.md)** or **[Reports](../../../api/Access.Reports.md)** collection, followed by the name of the property and its value. For example, to set the **Visible** property of the Customers form to **True** (-1), use the following line of code:

```vb
Forms!Customers.Visible = True
```

You can also set a property of a form or report from within the object's module by using the object's **Me** property. Code that uses the **Me** property executes faster than code that uses a fully qualified object name. For example, to set the **RecordSource** property of the Customers form to an SQL statement that returns all records with a CompanyName field entry beginning with "A" from within the Customers form module, use the following line of code:

```vb
Me.RecordSource = "SELECT * FROM Customers " _ 
    & "WHERE CompanyName Like 'A*'"
```


## Set a property of a control

Refer to the control in the **[Controls](../../../api/Access.Controls.md)** collection of the **Form** or **Report** object on which it resides. You can refer to the **Controls** collection either implicitly or explicitly, but the code executes faster if you use an implicit reference. The following examples set the **Visible** property of a text box called CustomerID on the Customers form:


```vb
' Faster method. 
Me!CustomerID.Visible = True
```


```vb
' Slower method. 
Forms!Customers.Controls!CustomerID.Visible = True
```

The fastest way to set a property of a control is from within an object's module by using the object's **Me** property. For example, you can use the following code to toggle the **Visible** property of a text box called CustomerID on the Customers form:




```vb
With Me!CustomerID 
    .Visible = Not .Visible 
End With
```


## Set a property of a form or report section

Refer to the form or report within the **Forms** or **Reports** collection, followed by the **Section** property and the integer or constant that identifies the section. The following examples set the **Visible** property of the page header section of the Customers form to **False**:


```vb
Forms!Customers.Section(3).Visible = False
```


```vb
Me!Section(acPageHeader).Visible = False
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]