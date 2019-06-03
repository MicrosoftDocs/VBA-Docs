---
title: UserDefinedProperty object (Outlook)
keywords: vbaol11.chm3151
f1_keywords:
- vbaol11.chm3151
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty
ms.assetid: aebe38db-0ff9-79d2-b5a7-751fea7c97f3
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperty object (Outlook)

Represents the definition of a user-defined property for a **[Folder](Outlook.Folder.md)** object.


## Remarks

Use **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** (_index_), where _index_ is a name or index number, to return a single **UserDefinedProperty** object.

Use the **[Add](Outlook.UserDefinedProperties.Add.md)** method of the **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** collection for a **Folder** object to define a user-defined property for that folder.

Use the **[Type](Outlook.UserDefinedProperty.Type.md)** property to return the user-defined property type and the **[DisplayFormat](Outlook.UserDefinedProperty.DisplayFormat.md)** property to return the display format for the user-defined property. If the **Type** property is set to **olCombination** or **olFormula**, use the **[Formula](Outlook.UserDefinedProperty.Formula.md)** property to return the formula used to generate values for the user-defined property.

The **UserDefinedProperty** object represents only the definition of a user-defined property, which is applicable to all Outlook items contained by the folder. To retrieve or change user-defined property values for an Outlook item in that folder, use the **[UserProperties](Outlook.MailItem.UserProperties.md)** property of the Outlook item, such as a **[MailItem](Outlook.MailItem.md)** object, to retrieve the **[UserProperties](Outlook.UserProperties.md)** collection for that item. You can then use the **[UserProperty](Outlook.UserProperty.md)** object for the appropriate user-defined property to retrieve or change the value of that user-defined property for the Outlook item.


## Example

The following Visual Basic for Applications (VBA) example displays the name of a specified **Folder** object, as well as the name and type of every **UserDefinedProperty** object contained in the **UserDefinedProperties** collection of the specified **Folder** object, to the Immediate window.


```vb
Sub DisplayUserProperties(ByRef FolderToCheck As Folder) 
 Dim objProperty As UserDefinedProperty 
 
 ' Print the name of the specified Folder object 
 ' reference to the Immediate window. 
 Debug.Print "--- Folder: " & FolderToCheck.Name 
 
 ' Check if there are any user-defined properties 
 ' associated with the Folder object reference. 
 If FolderToCheck.UserDefinedProperties.Count = 0 Then 
 ' No user-defined properties are present. 
 Debug.Print " No user-defined properties." 
 Else 
 ' Iterate through every user-defined property in 
 ' the folder. 
 For Each objProperty In FolderToCheck.UserDefinedProperties 
 ' Retrieve the name of the user-defined property. 
 strPropertyInfo = objProperty.Name 
 ' Retrieve the type of the user-defined property. 
 Select Case objProperty.Type 
 Case OlUserPropertyType.olCombination 
 strPropertyInfo = strPropertyInfo & " (Combination)" 
 Case OlUserPropertyType.olCurrency 
 strPropertyInfo = strPropertyInfo & " (Currency)" 
 Case OlUserPropertyType.olDateTime 
 strPropertyInfo = strPropertyInfo & " (Date/Time)" 
 Case OlUserPropertyType.olDuration 
 strPropertyInfo = strPropertyInfo & " (Duration)" 
 Case OlUserPropertyType.olEnumeration 
 strPropertyInfo = strPropertyInfo & " (Enumeration)" 
 Case OlUserPropertyType.olFormula 
 strPropertyInfo = strPropertyInfo & " (Formula)" 
 Case OlUserPropertyType.olInteger 
 strPropertyInfo = strPropertyInfo & " (Integer)" 
 Case OlUserPropertyType.olKeywords 
 strPropertyInfo = strPropertyInfo & " (Keywords)" 
 Case OlUserPropertyType.olNumber 
 strPropertyInfo = strPropertyInfo & " (Number)" 
 Case OlUserPropertyType.olOutlookInternal 
 strPropertyInfo = strPropertyInfo & " (Outlook Internal)" 
 Case OlUserPropertyType.olPercent 
 strPropertyInfo = strPropertyInfo & " (Percent)" 
 Case OlUserPropertyType.olSmartFrom 
 strPropertyInfo = strPropertyInfo & " (Smart From)" 
 Case OlUserPropertyType.olText 
 strPropertyInfo = strPropertyInfo & " (Text)" 
 Case OlUserPropertyType.olYesNo 
 strPropertyInfo = strPropertyInfo & " (Yes/No)" 
 Case Else 
 strPropertyInfo = strPropertyInfo & " (Unknown)" 
 End Select 
 
 ' Print the name and type of the user-defined property 
 ' to the Immediate window. 
 Debug.Print strPropertyInfo 
 Next 
 End If 
End Sub 

```


## Methods



|Name|
|:-----|
|[Delete](Outlook.UserDefinedProperty.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.UserDefinedProperty.Application.md)|
|[Class](Outlook.UserDefinedProperty.Class.md)|
|[DisplayFormat](Outlook.UserDefinedProperty.DisplayFormat.md)|
|[Formula](Outlook.UserDefinedProperty.Formula.md)|
|[Name](Outlook.UserDefinedProperty.Name.md)|
|[Parent](Outlook.UserDefinedProperty.Parent.md)|
|[Session](Outlook.UserDefinedProperty.Session.md)|
|[Type](Outlook.UserDefinedProperty.Type.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]