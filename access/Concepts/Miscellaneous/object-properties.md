---
title: Object properties
ROBOTS: INDEX
keywords: vbaac10.chm5187733
f1_keywords:
- vbaac10.chm5187733
ms.prod: access
ms.assetid: 9fc87446-68bd-d592-71c8-8d8c022af2c4
ms.date: 06/08/2019
localization_priority: Normal
---


# Object properties

**Applies to:** Access 2013 | Access 2016

The **Object** properties provide general information about objects contained in the Navigation Pane.

> [!NOTE] 
> The **Object** properties are available for all objects in a Microsoft Access database and only for forms, macros, modules, and reports in an Access project (.adp).


## Setting

You can view the **Object** properties, and set the **Description** or **Attributes** properties, in the following ways:

- Click an object in the Database window. On the **Database Tools** tab, in the **Show/Hide** group, click **Property Sheet**.
    
- Right-click an object in the Database window, and then click **Properties** on the shortcut menu.
    
You can also specify or determine the **Object** properties in an Access database by using Visual Basic . The **Object** properties of an Access project (.adp) are not available using Visual Basic.

> [!NOTE] 
> You can only enter or edit the **Description** and **Attributes** properties. The other **Object** properties are set by Microsoft Access and are read-only.


## Remarks

The objects for which you can display properties in the Database window are tables, queries, forms, reports, macros, and modules. Each class of objects in the database is represented by a separate DAO **Document** object within the DAO **Containers** collection. For example, the **Containers** collection contains a **Document** object that represents all the forms in the database.

The following **Object** properties are available from the Database window.

|Property|Description|
|:-----|:-----|
|**Name**|This is the name of the object and contains the setting from the object's **Name** property.|
|**Type**|This is the object's type. Microsoft Access object types are Form, Macro, Module, Query, Report, and Table.|
|**Description**|This is the object's description and is the same as the setting for the object's **Description** property. You can also set the object's **Description** property in the object's property sheet.|
|**Created**|This is the date that the object was created. For tables and queries, this property is the same as the **DateCreated** property.|
|**Modified**|This is the date that the object was last modified. For tables and queries, this property is the same as the **LastUpdated** property.|
|**Owner**|This is the owner of the object. For more information, see the **Owner** property.|
|**Attributes**|This property specifies whether the object is hidden or visible and whether the object can be replicated in a database replica. If you set the Hidden attribute to **True** (by selecting the **Hidden** check box), the object won't appear in the Database window.<br/><br/>To display hidden objects in the Navigation Pane, click the **Microsoft Office Button**, and then click **Access Options**. Click the **Current Database** category, and then click **Navigation Options**. Click **Show Hidden Objects**, and then click **OK**.<br/><br/>The icons for hidden objects will be dimmed in the Database window. You can then turn the Hidden attribute off, making the objects visible in the Database window.|

## Example

The following example uses the PrintObjectProperties subroutine to print the values of an object's **Object** properties to the Debug window. The subroutine requires the object type and object name as arguments.


```vb
Dim strObjectType As String 
Dim strObjectName As String 
Dim strMsg As String 
 
strMsg = "Enter object type (e.g., Forms, Scripts, " _ 
 & "Modules, Reports, Tables)." 
' Get object type. 
strObjectType = InputBox(strMsg) 
strMsg = "Enter the name of a form, macro, module, " _ 
 & "query, report, or table." 
' Get object name from user. 
strObjectName = InputBox(strMsg) 
' Pass object type and object name to 
' PrintObjectProperties subroutine. 
PrintObjectProperties strObjectType, strObjectName 
 
Sub PrintObjectProperties(strObjectType As String, strObjectName _ 
 As String) 
Dim dbs As Database, ctr As Container, doc As Document 
Dim intI As Integer 
Dim strTabChar As String 
Dim prp As DAO.Property 
 
Set dbs = CurrentDb 
strTabChar = vbTab 
' Set Container object variable. 
Set ctr = dbs.Containers(strObjectType) 
' Set Document object variable. 
Set doc = ctr.Documents(strObjectName) 
doc.Properties.Refresh 
' Print the object name to Debug window. 
Debug.Print doc.Name 
' Print each Object property to Debug window. 
For Each prp in doc.Properties 
 Debug.Print strTabChar & prp.Name & " = " & prp.Value 
Next 
End Sub
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]