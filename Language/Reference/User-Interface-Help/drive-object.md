---
title: Drive object
keywords: vblr6.chm2181923
f1_keywords:
- vblr6.chm2181923
ms.prod: office
api_name:
- Office.Drive
ms.assetid: 95229345-790b-d77d-c3b4-6b4998aa0336
ms.date: 11/12/2018
localization_priority: Normal
---


# Drive object

Provides access to the properties of a particular disk drive or network share.

## Remarks

The following code illustrates the use of the **Drive** object to access drive properties.

```vb
Sub ShowFreeSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " & UCase(drvPath) & " - " 
    s = s & d.VolumeName  & vbCrLf
    s = s & "Free Space: " & FormatNumber(d.FreeSpace/1024, 0) 
    s = s & " Kbytes"
    MsgBox s
End Sub
```

## Collections

|Collection|Description|
|:---------|:----------|
|[Drives](drives-collection.md)|Read-only collection of all available drives. |

## Properties

|Property|Description|
|:-------|:----------|
|[AvailableSpace](availablespace-property.md)|Returns the amount of available space to a user on a specified drive or network share. |
|[DriveLetter](driveletter-property.md)|Returns one uppercase letter that identifies the local drive or a network share. |
|[DriveType](drivetype-property.md)|Returns the type of a specified drive. |
|[FileSystem](filesystem-property.md)|Returns the file system in use for a specified drive. |
|[FreeSpace](freespace-property.md)|Returns the amount of free space to a user on a specified drive or network share. |
|[IsReady](isready-property.md)|Returns true if the specified drive is ready, and false if not. |
|[Path](path-property-filesystemobject-object.md)|Returns an uppercase letter followed by a colon that indicates the path name for a specified drive. |
|[RootFolder](rootfolder-property.md)|Returns a **[Folder](folder-object.md)** object that represents the root folder of a specified drive. |
|[SerialNumber](serialnumber-property.md)|Returns the serial number of a specified drive. |
|[ShareName](sharename-property.md)|Returns the network share name for a specified drive. |
|[TotalSize](totalsize-property.md)|Returns the total size of a specified drive or network share. |
|[VolumeName](volumename-property.md)|Sets or returns the volume name of a specified drive. |

## See also

- [Drives property](drives-property.md)
- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]