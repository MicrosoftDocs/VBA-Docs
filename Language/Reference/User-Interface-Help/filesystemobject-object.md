---
title: FileSystemObject object
keywords: vblr6.chm2181927
f1_keywords:
- vblr6.chm2181927
ms.prod: office
api_name:
- Office.FileSystemObject
ms.assetid: 7ad2dad3-c6d8-90a6-77a5-c712da8316f3
ms.date: 11/12/2018
---


# FileSystemObject object

Provides access to a computer's file system.

## Syntax

**Scripting.FileSystemObject**

## Remarks

The following code illustrates how the **FileSystemObject** is used to return a **[TextStream](textstream-object.md)** object that can be read from or written to:

```vb
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.WriteLine("This is a test.")
a.Close
```

In the example code:

- The **[CreateObject](createobject-function.md)** function returns the **FileSystemObject** (`fs`). 
- The **[CreateTextFile](createtextfile-method.md)** method creates the file as a **[TextStream](textstream-object.md)** object (`a`).
- The **[WriteLine](writeline-method.md)** method writes a line of text to the created text file. 
- The **[Close](close-method-filesystemobject-object.md)** method flushes the buffer and closes the file.

## Methods

- [BuildPath](buildpath-method.md)
- [Close](close-method-filesystemobject-object.md)
- [CopyFile](copyfile-method.md)
- [CopyFolder](copyfolder-method.md)
- [CreateFolder](createfolder-method.md)
- [CreateTextFile](createtextfile-method.md)
- [DeleteFile](deletefile-method.md)
- [DeleteFolder](deletefolder-method.md)
- [DriveExists](driveexists-method.md)
- [FileExists](fileexists-method.md)
- [FolderExists](folderexists-method.md)
- [GetAbsolutePathName](getabsolutepathname-method.md)
- [GetBaseName](getbasename-method.md)
- [GetDrive](getdrive-method.md)
- [GetDriveName](getdrivename-method.md)
- [GetExtensionName](getextensionname-method.md)
- [GetFile](getfile-method.md)
- [GetFileName](getfilename-method-visual-basic-for-applications.md)
- [GetFolder](getfolder-method.md)
- [GetParentFolderName](getparentfoldername-method.md)
- [GetSpecialFolder](getspecialfolder-method.md)
- [GetTempName](gettempname-method.md)
- [Move](move-method-filesystemobject-object.md)
- [MoveFile](movefile-method.md)
- [MoveFolder](movefolder-method.md)
- [OpenAsTextStream](openastextstream-method.md)
- [OpenTextFile](opentextfile-method.md)
- [Remove](remove-method-filesystemobject-object.md)
- [WriteLine](writeline-method.md)

## Properties

- [Count](count-property-filesystemobject-object.md)
- [Drives](drives-property.md)
- [Item](item-property-filesystemobject-object.md)
- [Name](name-property-filesystemobject-object.md)
- [Path](path-property-filesystemobject-object.md)
- [Size](size-property-filesystemobject-object.md)
- [Type](type-property-filesystemobject-object.md)

## See also

- [Methods (VBA)](../methods-visual-basic-for-applications.md)
- [Properties (VBA)](../properties-visual-basic-for-applications.md)