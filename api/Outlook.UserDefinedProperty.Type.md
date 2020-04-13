---
title: UserDefinedProperty.Type property (Outlook)
keywords: vbaol11.chm7
f1_keywords:
- vbaol11.chm7
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty.Type
ms.assetid: 94895d2b-7b3e-e455-3b58-58abd8279c10
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperty.Type property (Outlook)

Returns an **[OlUserPropertyType](Outlook.OlUserPropertyType.md)** constant indicating the type of the **[UserDefinedProperty](Outlook.UserDefinedProperty.md)** object. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a [UserDefinedProperty](Outlook.UserDefinedProperty.md) object.


## Example

The following Visual Basic for Applications (VBA) example displays the name of a specified  **[Folder](Outlook.Folder.md)** object, as well as the name and type of every **UserDefinedProperty** object contained in the **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** collection of the specified **Folder** object, to the Immediate window.


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


## See also


[UserDefinedProperty Object](Outlook.UserDefinedProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]