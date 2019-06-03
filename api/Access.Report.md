---
title: Report object (Access)
keywords: vbaac10.chm13901
f1_keywords:
- vbaac10.chm13901
ms.prod: access
api_name:
- Access.Report
ms.assetid: 6f77c1b4-a9ce-7caa-204c-fe0755c6f9df
ms.date: 03/09/2019
localization_priority: Normal
---


# Report object (Access)

A **Report** object refers to a particular Microsoft Access report.


## Remarks

A **Report** object is a member of the **[Reports](access.reports.md)** collection, which is a collection of all currently open reports. Within the **Reports** collection, individual reports are indexed beginning with zero. You can refer to an individual **Report** object in the **Reports** collection either by referring to the report by name, or by referring to its index within the collection. If the report name includes a space, the name must be surrounded by brackets ([ ]).

|Syntax|Example|
|:-----|:------|
|**Reports**!_reportname_|Reports!OrderReport|
|**Reports**![_report name_]|Reports![Order Report]|
|**Reports**("_reportname_")|Reports("OrderReport")|
|**Reports**(_index_)|Reports(0)|

> [!NOTE]
> Each **Report** object has a **Controls** collection, which contains all controls on the report. You can refer to a control on a report either by implicitly or explicitly referring to the **Controls** collection. Your code will be faster if you refer to the **Controls** collection implicitly. The following examples show two of the ways you might refer to a control named **NewData** on a report called **OrderReport**. 

```vb
' Implicit reference. 
Reports!OrderReport!NewData
```


```vb
' Explicit reference. 
Reports!OrderReport.Controls!NewData
```


## Example

The following example shows how to use the **NoData** event of a report to prevent the report from opening when there is no data to be displayed.

```vb
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " & _
        "The report will now close."
    Cancel = True

End Sub
```

<br/>

The following example shows how to use the **Page** event to add a watermark to a report before it is printed.

```vb
Private Sub Report_Page()
    Dim strWatermarkText As String
    Dim sizeHor As Single
    Dim sizeVer As Single

#If RUN_PAGE_EVENT = True Then
    With Me
        '// Print page border
        Me.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbBlack, B
    
        '// Print watermark
        strWatermarkText = "Confidential"
        
        .ScaleMode = 3
        .FontName = "Segoe UI"
        .FontSize = 48
        .ForeColor = RGB(255, 0, 0)

        '// Calculate text metrics
        sizeHor = .TextWidth(strWatermarkText)
        sizeVer = .TextHeight(strWatermarkText)
        
        '// Set the print location
        .CurrentX = (.ScaleWidth / 2) - (sizeHor / 2)
        .CurrentY = (.ScaleHeight / 2) - (sizeVer / 2)
    
        '// Print the watermark
        .Print strWatermarkText
    End With
#End If

End Sub
```

<br/>

The following example shows how to set the **BackColor** property of a control based on its value.

```vb
Private Sub SetControlFormatting()
    If (Me.AvgOfRating >= 8) Then
        Me.AvgOfRating.BackColor = vbGreen
    ElseIf (Me.AvgOfRating >= 5) Then
        Me.AvgOfRating.BackColor = vbYellow
    Else
        Me.AvgOfRating.BackColor = vbRed
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub

Private Sub Detail_Paint()
    ' do conditional formatting for the control in report view
    SetControlFormatting
End Sub
```

<br/>

The following example shows how to format a report to show progress bars. The example uses a pair of rectangle controls, **boxInside** and **boxOutside**, to create a progress bar based on the value of **AvgOfRating**. The progress bars are visible only when the report is opened in **Print Preview** mode or it is printed.

```vb
Private Sub Report_Load()
    If (Me.CurrentView = AcCurrentView.acCurViewPreview) Then
        Me.boxInside.Visible = True
        Me.boxOutside.Visible = True
    Else
        Me.boxInside.Visible = False
        Me.boxOutside.Visible = False
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub
```


## Events

- [Activate](Access.Report.Activate.md)
- [ApplyFilter](Access.Report.ApplyFilter.md)
- [Click](Access.Report.Click.md)
- [Close](Access.Report.Close.md)
- [Current](Access.Report.Current.md)
- [DblClick](Access.Report.DblClick.md)
- [Deactivate](Access.Report.Deactivate.md)
- [Error](Access.Report.Error.md)
- [Filter](Access.Report.Filter(even).md)
- [GotFocus](Access.Report.GotFocus.md)
- [KeyDown](Access.Report.KeyDown.md)
- [KeyPress](Access.Report.KeyPress.md)
- [KeyUp](Access.Report.KeyUp.md)
- [Load](Access.Report.Load.md)
- [LostFocus](Access.Report.LostFocus.md)
- [MouseDown](Access.Report.MouseDown.md)
- [MouseMove](Access.Report.MouseMove.md)
- [MouseUp](Access.Report.MouseUp.md)
- [MouseWheel](Access.Report.MouseWheel(even).md)
- [NoData](Access.Report.NoData.md)
- [Open](Access.Report.Open.md)
- [Page](Access.Report.Page(even).md)
- [Resize](Access.Report.Resize.md)
- [Timer](Access.Report.Timer.md)
- [Unload](Access.Report.Unload.md)

## Methods

- [Circle](Access.Report.Circle.md)
- [Line](Access.Report.Line.md)
- [Move](Access.Report.Move.md)
- [Print](Access.Report.Print.md)
- [PSet](Access.Report.PSet.md)
- [Requery](Access.Report.Requery.md)
- [Scale](Access.Report.Scale.md)
- [TextHeight](Access.Report.TextHeight.md)
- [TextWidth](Access.Report.TextWidth.md)

## Properties

- [ActiveControl](Access.Report.ActiveControl.md)
- [AllowLayoutView](Access.Report.AllowLayoutView.md)
- [AllowReportView](Access.Report.AllowReportView.md)
- [Application](Access.Report.Application.md)
- [AutoCenter](Access.Report.AutoCenter.md)
- [AutoResize](Access.Report.AutoResize.md)
- [BorderStyle](Access.Report.BorderStyle.md)
- [Caption](Access.Report.Caption.md)
- [CloseButton](Access.Report.CloseButton.md)
- [ControlBox](Access.Report.ControlBox.md)
- [Controls](Access.Report.Controls.md)
- [Count](Access.Report.Count.md)
- [CurrentRecord](Access.Report.CurrentRecord.md)
- [CurrentView](Access.Report.CurrentView.md)
- [CurrentX](Access.Report.CurrentX.md)
- [CurrentY](Access.Report.CurrentY.md)
- [Cycle](Access.Report.Cycle.md)
- [DateGrouping](Access.Report.DateGrouping.md)
- [DefaultControl](Access.Report.DefaultControl.md)
- [DefaultView](Access.Report.DefaultView.md)
- [Dirty](Access.Report.Dirty.md)
- [DisplayOnSharePointSite](Access.Report.DisplayOnSharePointSite.md)
- [DrawMode](Access.Report.DrawMode.md)
- [DrawStyle](Access.Report.DrawStyle.md)
- [DrawWidth](Access.Report.DrawWidth.md)
- [FastLaserPrinting](Access.Report.FastLaserPrinting.md)
- [FillColor](Access.Report.FillColor.md)
- [FillStyle](Access.Report.FillStyle.md)
- [Filter](Access.Report.Filter(property).md)
- [FilterOn](Access.Report.FilterOn.md)
- [FilterOnLoad](Access.Report.FilterOnLoad.md)
- [FitToPage](Access.Report.FitToPage.md)
- [FontBold](Access.Report.FontBold.md)
- [FontItalic](Access.Report.FontItalic.md)
- [FontName](Access.Report.FontName.md)
- [FontSize](Access.Report.FontSize.md)
- [FontUnderline](Access.Report.FontUnderline.md)
- [ForeColor](Access.Report.ForeColor.md)
- [FormatCount](Access.Report.FormatCount.md)
- [GridX](Access.Report.GridX.md)
- [GridY](Access.Report.GridY.md)
- [GroupLevel](Access.Report.GroupLevel.md)
- [GrpKeepTogether](Access.Report.GrpKeepTogether.md)
- [HasData](Access.Report.HasData.md)
- [HasModule](Access.Report.HasModule.md)
- [Height](Access.Report.Height.md)
- [HelpContextId](Access.Report.HelpContextId.md)
- [HelpFile](Access.Report.HelpFile.md)
- [Hwnd](Access.Report.Hwnd.md)
- [InputParameters](Access.Report.InputParameters.md)
- [KeyPreview](Access.Report.KeyPreview.md)
- [LayoutForPrint](Access.Report.LayoutForPrint.md)
- [Left](Access.Report.Left.md)
- [MenuBar](Access.Report.MenuBar.md)
- [MinMaxButtons](Access.Report.MinMaxButtons.md)
- [Modal](Access.Report.Modal.md)
- [Module](Access.Report.Module.md)
- [MouseWheel](Access.Report.MouseWheel(property).md)
- [Moveable](Access.Report.Moveable.md)
- [MoveLayout](Access.Report.MoveLayout.md)
- [Name](Access.Report.Name.md)
- [NextRecord](Access.Report.NextRecord.md)
- [OnActivate](Access.Report.OnActivate.md)
- [OnApplyFilter](Access.Report.OnApplyFilter.md)
- [OnClick](Access.Report.OnClick.md)
- [OnClose](Access.Report.OnClose.md)
- [OnCurrent](Access.Report.OnCurrent.md)
- [OnDblClick](Access.Report.OnDblClick.md)
- [OnDeactivate](Access.Report.OnDeactivate.md)
- [OnError](Access.Report.OnError.md)
- [OnFilter](Access.Report.OnFilter.md)
- [OnGotFocus](Access.Report.OnGotFocus.md)
- [OnKeyDown](Access.Report.OnKeyDown.md)
- [OnKeyPress](Access.Report.OnKeyPress.md)
- [OnKeyUp](Access.Report.OnKeyUp.md)
- [OnLoad](Access.Report.OnLoad.md)
- [OnLostFocus](Access.Report.OnLostFocus.md)
- [OnMouseDown](Access.Report.OnMouseDown.md)
- [OnMouseMove](Access.Report.OnMouseMove.md)
- [OnMouseUp](Access.Report.OnMouseUp.md)
- [OnNoData](Access.Report.OnNoData.md)
- [OnOpen](Access.Report.OnOpen.md)
- [OnPage](Access.Report.OnPage.md)
- [OnResize](Access.Report.OnResize.md)
- [OnTimer](Access.Report.OnTimer.md)
- [OnUnload](Access.Report.OnUnload.md)
- [OpenArgs](Access.Report.OpenArgs.md)
- [OrderBy](Access.Report.OrderBy.md)
- [OrderByOn](Access.Report.OrderByOn.md)
- [OrderByOnLoad](Access.Report.OrderByOnLoad.md)
- [Orientation](Access.Report.Orientation.md)
- [Page](Access.Report.Page(property).md)
- [PageFooter](Access.Report.PageFooter.md)
- [PageHeader](Access.Report.PageHeader.md)
- [Pages](Access.Report.Pages.md)
- [Painting](Access.Report.Painting.md)
- [PaintPalette](Access.Report.PaintPalette.md)
- [PaletteSource](Access.Report.PaletteSource.md)
- [Parent](Access.Report.Parent.md)
- [Picture](Access.Report.Picture.md)
- [PictureAlignment](Access.Report.PictureAlignment.md)
- [PictureData](Access.Report.PictureData.md)
- [PicturePages](Access.Report.PicturePages.md)
- [PicturePalette](Access.Report.PicturePalette.md)
- [PictureSizeMode](Access.Report.PictureSizeMode.md)
- [PictureTiling](Access.Report.PictureTiling.md)
- [PictureType](Access.Report.PictureType.md)
- [PopUp](Access.Report.PopUp.md)
- [PrintCount](Access.Report.PrintCount.md)
- [Printer](Access.Report.Printer.md)
- [PrintSection](Access.Report.PrintSection.md)
- [Properties](Access.Report.Properties.md)
- [PrtDevMode](Access.Report.PrtDevMode.md)
- [PrtDevNames](Access.Report.PrtDevNames.md)
- [PrtMip](Access.Report.PrtMip.md)
- [RecordLocks](Access.Report.RecordLocks.md)
- [Recordset](Access.Report.Recordset.md)
- [RecordSource](Access.Report.RecordSource.md)
- [RecordSourceQualifier](Access.Report.RecordSourceQualifier.md)
- [Report](Access.Report.Report.md)
- [RibbonName](Access.Report.RibbonName.md)
- [ScaleHeight](Access.Report.ScaleHeight.md)
- [ScaleLeft](Access.Report.ScaleLeft.md)
- [ScaleMode](Access.Report.ScaleMode.md)
- [ScaleTop](Access.Report.ScaleTop.md)
- [ScaleWidth](Access.Report.ScaleWidth.md)
- [ScrollBars](Access.Report.ScrollBars.md)
- [Section](Access.Report.Section.md)
- [ServerFilter](Access.Report.ServerFilter.md)
- [Shape](Access.report.shape.md)
- [ShortcutMenuBar](Access.Report.ShortcutMenuBar.md)
- [ShowPageMargins](Access.Report.ShowPageMargins.md)
- [Tag](Access.Report.Tag.md)
- [TimerInterval](Access.Report.TimerInterval.md)
- [Toolbar](Access.Report.Toolbar.md)
- [Top](Access.Report.Top.md)
- [UseDefaultPrinter](Access.Report.UseDefaultPrinter.md)
- [Visible](Access.Report.Visible.md)
- [Width](Access.Report.Width.md)
- [WindowHeight](Access.Report.WindowHeight.md)
- [WindowLeft](Access.Report.WindowLeft.md)
- [WindowTop](Access.Report.WindowTop.md)
- [WindowWidth](Access.Report.WindowWidth.md)



## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
