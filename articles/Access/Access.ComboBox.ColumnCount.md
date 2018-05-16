---
title: ComboBox.ColumnCount Property (Access)
keywords: vbaac10.chm11380
f1_keywords:
- vbaac10.chm11380
ms.prod: access
api_name:
- Access.ComboBox.ColumnCount
ms.assetid: 76db2415-ee22-89c6-6753-f20d636d41f8
ms.date: 06/08/2017
---


# ComboBox.ColumnCount Property (Access)

You can use the  **ColumnCount** property to specify the number of columns displayed in a list box or in the list box portion of a combo box, or sent to OLE objects in a chart control or unbound object frame . Read/write **Integer**.


## Syntax

 _expression_. **ColumnCount**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

The  **ColumnCount** property holds an integer between 1 and the maximum number of fields in the table, query, or SQL statement, or the maximum number of values in the value list, specified in the **RowSource** property of the control.

For [table fields](table-field.md) , you can set this property on the **Lookup** tab in the Field Properties section of table Design view for fields with the **DisplayControl** property set to Combo Box or List Box.

For example, if you set the  **ColumnCount** property for a list box on an Employees form to 3, one column can list last names, another can list first names, and the third can list employee ID numbers.

A combo box or list box can have multiple columns. If the control's  **RowSource** property contains the name of a table, query, or SQL statement, a combo box or list box will display the fields from that source, from left to right, up to the number specified by the **ColumnCount** property.

To display a different combination of fields, create either a new query or a new SQL statement for the  **RowSource** property, specifying the fields and the order you want.

If the  **RowSource** property contains a list of values (the **RowSourceType** property is set to Value List), the values are put into the rows and columns of the combo box or list box in the order they are listed in the **RowSource** property. For example, if the **RowSource** property contains the list "Red; Green; Blue; Yellow" and the **ColumnCount** property is set to 2, the first row of the combo box or list box list will contain "Red" in the first column and "Green" in the second column. The second row will contain "Blue" in the first column and "Yellow" in the second column.

You can use the  **ColumnWidths** property to set the width of the columns displayed in the control, or to hide columns.


## Example

The following example uses the  **Column** property and the **ColumnCount** property to print the values of a list box selection.


```vb
Public Sub Read_ListBox() 
 
    Dim intNumColumns As Integer 
    Dim intI As Integer 
    Dim frmCust As Form 
 
    Set frmCust = Forms!frmCustomers 
    If frmCust!lstCustomerNames.ItemsSelected.Count > 0 Then 
 
        ' Any selection? 
        intNumColumns = frmCust!lstCustomerNames.ColumnCount 
        Debug.Print "The list box contains "; intNumColumns; _ 
            IIf(intNumColumns = 1, " column", " columns"); _ 
             " of data." 
 
        Debug.Print "The current selection contains:" 
        For intI = 0 To intNumColumns - 1 
            ' Print column data. 
            Debug.Print frmCust!lstCustomerNames.Column(intI) 
        Next intI 
    Else 
        Debug.Print "You haven't selected an entry in the " _ 
            &; "list box." 
    End If 
 
    Set frmCust = Nothing 
 
End Sub
```



The following example show how to create a combo box that is bound to one column while displaying another. Setting the  **ColumnCount** property to 2 specifies that the **cboDept** combo box will display the first two columns of the data source specified by the **RowSource** property. Setting the **BoundColumn** property to 1 specifies that the value stored in the first column will be returned when you inspect the value of the combo box.

The  **ColumnWidths** property specifies the width of the two columns. By setting the width of the first column to **0in.**, the first column is not displayed in the combo box.

 **Sample code provided by:**
![MVP Contributor](images/odc_OfficeTA_33px_MVPContrib.jpg) Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)




```vb
Private Sub cboDept_Enter()
    With cboDept
        .RowSource = "SELECT * FROM tblDepartments ORDER BY Department"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in.;1in."
    End With
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](Access.ComboBox.md)

