---
title: Date and Time Criteria from a Control on a Form
keywords: vbaac10.chm5187147
f1_keywords:
- vbaac10.chm5187147
ms.prod: access
ms.assetid: 132cf297-eb3b-68a1-7a47-b832931a5091
ms.date: 06/08/2017
---


# Date and Time Criteria from a Control on a Form

If you want to change the  _criteria_ argument for an operation based on a user's decision, you can specify that the criteria comes from a control on a form. For example, you could specify that the _criteria_ argument comes from a list box containing order dates from an Orders table.

To specify date and time criteria that comes from a control on a form, you include in the  _criteria_ argument an expression that references the control on the form. This expression should be separate from the string expression, so that Access will evaluate the control expression first and concatenate it with the rest of the string expression before performing the appropriate operation.

In addition to enclosing the entire string expression in double quotation marks ("), you must also ensure that the date or time criteria within the string expression is enclosed in number signs (#). The number signs must be included in the strings flanking the expression that references the control on the form.


 **Note**  The number signs indicate to Access that the  _criteria_ argument contains a date or time within a string.

The following examples set a form's  **Filter** or **ServerFilter** property based on criteria that comes from a control named HireDate that's on the form. Note the placement of the number signs.



```vb
Forms!Employees.Filter = "[HireDate] >= #" _ 
 &; Forms!Employees!HireDate &; "#" 
Forms!Employees.FilterOn = True
```

- or -



```vb
Forms!Employees.ServerFilter = "[HireDate] >= #" _ 
 &; Forms!Employees!HireDate &; "#" 
Forms!Employees.FilterOn = True
```

If the current value of the HireDate control is , the  **Filter** or **ServerFilter** property will have the following _criteria_ argument:



```text
"[HireDate] >= #5-1-92#"
```


 **Tip**   To troubleshoot an expression in the _criteria_ argument, break the expression into smaller components and test each individually in the Immediate window. When all of the components are working correctly, put them back together one at a time until the complete expression works correctly.

You can also include a variable representing a date or time in the  _criteria_ argument. The variable should be separate from the string expression, so that Access will evaluate the variable first and then concatenate it with the rest of the string expression. The date or time criteria must be enclosed in number signs.
The following example shows how to construct a  _criteria_ argument that includes a variable representing a date or time:



```vb
Dim datHireDate As Date 
datHireDate = #5-1-92# 
Forms!Employees.Filter = "[HireDate] >= #" _ 
 &; datHireDate &; "#"
```


