---
title: Application.OpenXML method (Project)
ms.prod: project-server
api_name:
- Project.Application.OpenXML
ms.assetid: dcf3dd0e-78ec-b95c-b890-dca5507acd92
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OpenXML method (Project)

Opens a project from an XML string.


## Syntax

_expression_. `OpenXML`( `_XML_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XML_|Required|**String**|String containing a valid Project XML string that conforms to the Project XML schema.|

## Return value

 **Long**


## Remarks

The Project XML schema is available in the Project SDK, as the file mspdi_pj15.xsd. You can create an XML file by saving a project to XML, and then editing the file. If you programmatically create an XML string, you should validate it against the schema before using it with the **OpenXML** method.

The **OpenXML** method returns 0 if it is successful.


> [!NOTE] 
> You can also use the **[FileOpenEx](Project.Application.FileOpenEx.md)** method to open a valid Project XML file. The **OpenXML** method is primarily designed to open a project by using an XML string.


## Example

The following example opens a file named OneTaskEdited.xml that was created by saving a project as XML and then editing the file to remove default values. The example requires a reference to the Microsoft Scripting Runtime library (scrrun.dll).


```vb
Sub ImportXMLProject() 
    ' Requires reference to the Microsoft Scripting Runtime library (scrrun.dll). 
    Dim txtStream As TextStream 
    Dim fileName As String 
    Dim xmlContents As String 
    Dim fsObject As FileSystemObject 
 
    fileName = "C:\Project\VBA\Samples\OneTaskEdited.xml" 
    Set fsObject = CreateObject("Scripting.FileSystemObject") 
 
    If Not fsObject.FileExists(fileName) Then 
        MsgBox "The file does not exist: " & vbCrLf & fileName 
    Else 
        ' Open a text stream. 
        Set txtStream = fsObject.OpenTextFile(fileName:=fileName, IOMode:=ForReading) 
 
        xmlContents = txtStream.ReadAll 
        Application.OpenXML(xmlContents) 
        txtStream.Close 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]