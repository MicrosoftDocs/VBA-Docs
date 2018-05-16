---
title: Form.DatasheetFontWeight Property (Access)
keywords: vbaac10.chm13398
f1_keywords:
- vbaac10.chm13398
ms.prod: access
api_name:
- Access.Form.DatasheetFontWeight
ms.assetid: 6dd2c6d3-1f27-8b86-abf5-f5581fbe7d23
ms.date: 06/08/2017
---


# Form.DatasheetFontWeight Property (Access)

You can use the  **DatasheetFontWeight** property to specify the line width of the font used to display and print characters for field names and data in Datasheet view. Read/write **Integer**.


## Syntax

 _expression_. **DatasheetFontWeight**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **DatasheetFontWeight** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

These properties are only available in [Visual Basic](set-properties-by-using-visual-basic.md) within a Microsoft Access database.

In Visual Basic, the  **DatasheetFontWeight** property setting uses the following **Integer** values.



|**Setting**|**Description**|
|:-----|:-----|
|100|Thin|
|200|Extra Light|
|300|Light|
|400|(Default) Normal|
|500|Medium|
|600|Semi-bold|
|700|Bold|
|800|Extra Bold|
|900|Heavy|
The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetForeColor](Access.Form.DatasheetForeColor.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**[DatasheetBackColor](Access.Form.DatasheetBackColor.md)**|
|**[DatasheetFontName](Access.Form.DatasheetFontName.md)** *|**[DatasheetGridlinesColor](Access.Form.DatasheetGridlinesColor.md)**|
|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|**[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)**|
|**DatasheetFontWeight** *|**[DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection of the database.


## Example

The following example sets the font to MS Serif, the font size to 10 points, and the font weight to medium (500) in Datasheet view of the Products table.


```vb
Sub SetDatasheetFont 
 
   Dim dbs As Object, objProducts As Object 
   Set dbs = CurrentDb 
   Const DB_Text As Long = 10 
   Const DB_Integer As Long = 3 
   Set objProducts = dbs!Products 
    
   SetTableProperty objProducts, "DatasheetFontName", DB_Text, "MS Serif" 
   SetTableProperty objProducts, "DatasheetFontHeight", DB_Integer, 10 
   SetTableProperty objProducts, "DatasheetFontWeight", DB_Integer, 500 
 
End Sub 
 
Sub SetTableProperty(objTableObj As Object, strPropertyName As String, _ 
        intPropertyType As Integer, varPropertyValue As Variant) 
    ' Set Microsoft Access-defined table property without causing 
    ' nonrecoverable run-time error. 
    Const conErrPropertyNotFound = 3270 
    Dim prpProperty As Variant 
    On Error Resume Next                ' Don't trap errors. 
    objTableObj.Properties(strPropertyName) = varPropertyValue 
    If Err <> 0 Then                    ' Error occurred when value set. 
        If Err <> conErrPropertyNotFound Then 
            On Error GoTo 0 
            MsgBox "Couldn't set property '" &; strPropertyName _ 
                &; "' on table '" &; objTableObj.Name &; "'", 48, "SetTableProperty" 
        Else 
            On Error GoTo 0 
            Set prpProperty = objTableObj.CreateProperty(strPropertyName, _ 
                intPropertyType, varPropertyValue) 
            objTableObj.Properties.Append prpProperty 
        End If 
    End If 
    objTableObj.Properties.Refresh 
End Sub
```

The next example makes the same changes as the preceding example in Datasheet view of the open Products form.




```vb
Forms!Products.DatasheetFontName = "MS Serif" 
Forms!Products.DatasheetFontHeight = 10 
Forms!Products.DatasheetFontWeight = 500
```


## See also


#### Concepts


[Form Object](Access.Form.md)

