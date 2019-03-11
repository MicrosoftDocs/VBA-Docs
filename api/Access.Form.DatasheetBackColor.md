---
title: Form.DatasheetBackColor property (Access)
keywords: vbaac10.chm13407
f1_keywords:
- vbaac10.chm13407
ms.prod: access
api_name:
- Access.Form.DatasheetBackColor
ms.assetid: 69734522-e570-86a5-f971-ce26ee4f88c3
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DatasheetBackColor property (Access)

You can use the **DatasheetBackColor** property in [Visual Basic](../access/Concepts/Settings/set-properties-by-using-visual-basic.md) to specify or determine the background color of an entire table, query, or form in Datasheet view within a Microsoft Access database. Read/write **Long**.


## Syntax

_expression_.**DatasheetBackColor**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The following setting information applies to both Microsoft Access database and Access projects (.adp).

Setting the **DatasheetBackColor** property for a table or query won't affect this property setting for a form that uses the table or query as its source of data.

The following table contains the properties that don't exist in the DAO **Properties** collection until you set them by using the **Formatting (Datasheet)** toolbar, or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.

|||
|:-----|:-----|
|**DatasheetBackColor**|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|
|**[DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)**|**[DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**DatasheetForeColor** *|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)**|
|**[DatasheetFontName](Access.Form.DatasheetFontName.md)** *|**[DatasheetGridlinesColor](Access.Form.DatasheetGridlinesBehavior.md)**|

> [!NOTE] 
> When you add or set any property listed with an asterisk, Access automatically adds it to the **Properties** collection.


## Example

The following example uses the **SetTableProperty** procedure to set a table's font color to dark blue and its background color to light gray. If a "Property not found" error occurs when the property is set, the **CreateProperty** method is used to add the property to the object's **Properties** collection.

```vb
Dim dbs As Object, objProducts As Object 
Const lngForeColor As Long = 8388608 ' Dark blue. 
Const lngBackColor As Long = 12632256 ' Light gray. 
Const DB_Long As Long = 4 
Set dbs = CurrentDb 
Set objProducts = dbs!Products 
SetTableProperty objProducts, "DatasheetBackColor", DB_Long, lngBackColor 
SetTableProperty objProducts, "DatasheetForeColor", DB_Long, lngForeColor 
 
Sub SetTableProperty(objTableObj As Object, strPropertyName As String, _ 
 intPropertyType As Integer, varPropertyValue As Variant) 
 Const conErrPropertyNotFound = 3270 
 Dim prpProperty As Variant 
 On Error Resume Next ' Don't trap errors. 
 objTableObj.Properties(strPropertyName) = varPropertyValue 
 If Err <> 0 Then ' Error occurred when value set. 
 If Err <> conErrPropertyNotFound Then 
 ' Error is unknown. 
 MsgBox "Couldn't set property '" & strPropertyName _ 
 & "' on table '" & tdfTableObj.Name & "'", vbExclamation, Err.Description 
 Err.Clear 
 Else 
 ' Error is "Property not found", so add it to collection. 
 Set prpProperty = objTableObj.CreateProperty(strPropertyName, _ 
 intPropertyType, varPropertyValue) 
 objTableObj.Properties.Append prpProperty 
 Err.Clear 
 End If 
 End If 
 objTableObj.Properties.Refresh 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]