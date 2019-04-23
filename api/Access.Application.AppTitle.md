---
title: Application.AppTitle property (Access)
keywords: vbaac10.chm5187013
f1_keywords:
- vbaac10.chm5187013
ms.prod: access
api_name:
- Access.Application. AppTitle
ms.assetid: a505f465-7813-6677-dd80-21a757c9d422
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.AppTitle property (Access)

You can use the **AppTitle** property to specify the text that appears in the application database's title bar. For example, you can use the **AppTitle** property to specify that the string "Inventory Control" appear in the title bar of your Inventory Control database application.

## Syntax

_expression_.**AppTitle**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.

## Setting

The **AppTitle** property is a string expression containing the text to appear in the title bar.

The easiest way to set this property is by using the **Application Title** option in the **Access Options** dialog box. You can also set this property by using a macro or Visual Basic.

To set the **AppTitle** property by using a macro or Visual Basic, you must first either set the property in the **Access Options** dialog box once or create the property in the following ways:

- In a Microsoft Access database, you can add it by using the **CreateProperty** method and append it to the **Properties** collection of the **Database** object.
    
- In a Microsoft Access project (.adp), you can add it to the **AccessObjectProperties** collection of the **[CurrentProject](access.currentproject.md)** object by using the **Add** method.
    
You must also use the **[RefreshTitleBar](access.application.refreshtitlebar.md)** method to make any changes visible immediately.


## Remarks

If this property isn't set, the string "Microsoft Access" appears in the title bar.

This property setting takes effect immediately after it is set in code (as long as the code includes the **RefreshTitleBar** method), or the **Access Options** dialog box is closed.

## Example

The following example shows how to change the **[AppIcon](Access.Application.AppIcon.md)** and **AppTitle** properties in a Microsoft Access database. If the properties haven't already been set or created, you must create them and append them to the **Properties** collection by using the **CreateProperty** method.


```vb
Sub cmdAddProp_Click() 
 Dim intX As Integer 
 Const DB_Text As Long = 10 
 intX = AddAppProperty("AppTitle", DB_Text, "My Custom Application") 
 intX = AddAppProperty("AppIcon", DB_Text, "C:\Windows\Cars.bmp") 
 CurrentDb.Properties("UseAppIconForFrmRpt") = 1 
 Application.RefreshTitleBar 
End Sub 
 
Function AddAppProperty(strName As String, _ 
 varType As Variant, varValue As Variant) As Integer 
 Dim dbs As Object, prp As Variant 
 Const conPropNotFoundError = 3270 
 
 Set dbs = CurrentDb 
 On Error GoTo AddProp_Err 
 dbs.Properties(strName) = varValue 
 AddAppProperty = True 
 
AddProp_Bye: 
 Exit Function 
 
AddProp_Err: 
 If Err = conPropNotFoundError Then 
 Set prp = dbs.CreateProperty(strName, varType, varValue) 
 dbs.Properties.Append prp 
 Resume 
 Else 
 AddAppProperty = False 
 Resume AddProp_Bye 
 End If 
End Function
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]