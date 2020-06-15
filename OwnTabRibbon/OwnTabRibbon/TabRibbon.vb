'https://www.cadcenter.pl/
'Kamil Martenczuk Autodesk Authorized Developer DEPL2710
'info@cadcenter.pl
'Date: 6/15/2020
'OwnTabRibbon
'**********************************************************
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Media.Imaging
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.ApplicationServices

Public Class TabBuilder

    Public Shared Sub CadCenter_Tab()
        Dim ribbonControl As Autodesk.Windows.RibbonControl = Autodesk.Windows.ComponentManager.Ribbon
        Dim Tab As RibbonTab = New RibbonTab()
        Tab.Title = "CadCenter"
        Tab.Id = "CadCenter"
        ribbonControl.Tabs.Add(Tab)
        '////////////////////////////////// P A N E L 1 /////////////////////////////////////
        Dim panel1Panel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
        panel1Panel.Title = "2D Tools"
        Dim Panel1 As RibbonPanel = New RibbonPanel()
        Panel1.Source = panel1Panel
        Tab.Panels.Add(Panel1)
        Dim pan1button0 As RibbonButton = New RibbonButton()
        pan1button0.ShowText = True
        pan1button0.ShowImage = True
        pan1button0.Orientation = System.Windows.Controls.Orientation.Vertical
        pan1button0.Size = RibbonItemSize.Large
        pan1button0.CommandHandler = New RibbonCommandHandler()
        Dim pan1button01 As RibbonButton = New RibbonButton()
        pan1button01 = pan1button0.Clone
        pan1button01.LargeImage = Images.getBitmap(My.Resources.large1)
        pan1button01.Text = "PolyLine"
        pan1button01.Name = "Polilinia"
        pan1button01.Tag = "_pline"
        Dim pan1button As RibbonButton = New RibbonButton()
        pan1button.ShowText = False
        pan1button.ShowImage = True
        pan1button.Image = Images.getBitmap(My.Resources.small)
        pan1button.LargeImage = Images.getBitmap(My.Resources.small)
        pan1button.CommandHandler = New RibbonCommandHandler()
        Dim pan1button1 As RibbonButton = New RibbonButton()
        pan1button1 = pan1button.Clone
        pan1button1.Tag = "_pline"
        Dim pan1button2 As RibbonButton = New RibbonButton()
        pan1button2 = pan1button.Clone
        pan1button2.Tag = "_pline"
        Dim pan1button3 As RibbonButton = New RibbonButton()
        pan1button3 = pan1button.Clone
        pan1button3.Tag = "_pline"
        Dim pan1button4 As RibbonButton = New RibbonButton()
        pan1button4 = pan1button.Clone
        pan1button4.Tag = "_pline"
        Dim pan1button5 As RibbonButton = New RibbonButton()
        pan1button5 = pan1button.Clone
        pan1button5.Tag = "_pline"
        Dim pan1button6 As RibbonButton = New RibbonButton()
        pan1button6 = pan1button.Clone
        pan1button6.Tag = "_pline"
        Dim pan1row1 As RibbonRowPanel = New RibbonRowPanel()
        pan1row1.Items.Add(pan1button1)
        pan1row1.Items.Add(pan1button2)
        pan1row1.Items.Add(New RibbonRowBreak())
        pan1row1.Items.Add(pan1button3)
        pan1row1.Items.Add(pan1button4)
        pan1row1.Items.Add(New RibbonRowBreak())
        pan1row1.Items.Add(pan1button5)
        pan1row1.Items.Add(pan1button6)
        pan1row1.Items.Add(New RibbonRowBreak())
        panel1Panel.Items.Add(pan1button01)
        panel1Panel.Items.Add(New RibbonSeparator())
        panel1Panel.Items.Add(pan1row1)
        '////////////////////////////////// P A N E L 2 /////////////////////////////////////
        Dim panel2Panel As RibbonPanelSource = New RibbonPanelSource()
        panel2Panel.Title = "3D Tools"
        Dim panel2 As RibbonPanel = New RibbonPanel()
        panel2.Source = panel2Panel
        Tab.Panels.Add(panel2)
        Dim pan2button As RibbonButton = New RibbonButton()
        pan2button = pan1button0.Clone
        Dim pan2button1 As RibbonButton = New RibbonButton()
        pan2button1 = pan2button.Clone
        pan2button1.Text = "Floor"
        pan2button1.LargeImage = Images.getBitmap(My.Resources.large2)
        pan2button1.Tag = "_pline"
        Dim pan2button2 As RibbonButton = New RibbonButton()
        pan2button2 = pan2button.Clone
        pan2button2.Text = "Walls"
        pan2button2.LargeImage = Images.getBitmap(My.Resources.large3)
        pan2button2.Tag = "_pline"
        Dim pan2button3 As RibbonButton = New RibbonButton()
        pan2button3 = pan2button.Clone
        pan2button3.Text = "Exterior" + vbCrLf + "Doors"
        pan2button3.LargeImage = Images.getBitmap(My.Resources.large4)
        pan2button3.Tag = "_pline"
        Dim pan2button4 As RibbonButton = New RibbonButton()
        pan2button4 = pan2button.Clone
        pan2button4.Text = "Interior" + vbCrLf + "Doors"
        pan2button4.LargeImage = Images.getBitmap(My.Resources.large5)
        pan2button4.Tag = "_pline"
        Dim pan2button5 As RibbonButton = New RibbonButton()
        pan2button5 = pan2button.Clone
        pan2button5.Text = "Glass" + vbCrLf + "Doors"
        pan2button5.LargeImage = Images.getBitmap(My.Resources.large6)
        pan2button5.Tag = "_pline"
        Dim pan2button6 As RibbonButton = New RibbonButton()
        pan2button6 = pan2button.Clone
        pan2button6.Text = "Windows"
        pan2button6.LargeImage = Images.getBitmap(My.Resources.large7)
        pan2button6.Tag = "_pline"
        Dim pan2button7 As RibbonButton = New RibbonButton()
        pan2button7 = pan2button.Clone
        pan2button7.Text = "Stairs"
        pan2button7.LargeImage = Images.getBitmap(My.Resources.large8)
        pan2button7.Tag = "_pline"
        Dim pan2button8 As RibbonButton = New RibbonButton()
        pan2button8 = pan2button.Clone
        pan2button8.Text = "Roof"
        pan2button8.LargeImage = Images.getBitmap(My.Resources.large9)
        pan2button8.Tag = "_pline"
        panel2Panel.Items.Add(pan2button1)
        panel2Panel.Items.Add(pan2button2)
        panel2Panel.Items.Add(New RibbonSeparator())
        panel2Panel.Items.Add(pan2button3)
        panel2Panel.Items.Add(pan2button4)
        panel2Panel.Items.Add(pan2button5)
        panel2Panel.Items.Add(New RibbonSeparator())
        panel2Panel.Items.Add(pan2button6)
        panel2Panel.Items.Add(New RibbonSeparator())
        panel2Panel.Items.Add(pan2button7)
        panel2Panel.Items.Add(New RibbonSeparator())
        panel2Panel.Items.Add(pan2button8)
        '////////////////////////////////// P A N E L 3 /////////////////////////////////////
        Dim panel3Panel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
        panel3Panel.Title = "Vizualisation"
        Dim Panel3 As RibbonPanel = New RibbonPanel()
        Panel3.Source = panel3Panel
        Tab.Panels.Add(Panel3)
        Dim pan3button As RibbonButton = New RibbonButton()
        pan3button = pan1button0.Clone
        Dim pan3button1 As RibbonButton = New RibbonButton()
        pan3button1 = pan3button.Clone
        pan3button1.Text = "Camera"
        pan3button1.LargeImage = Images.getBitmap(My.Resources.large10)
        pan3button1.Tag = "_pline"
        Dim pan3button2 As RibbonButton = New RibbonButton()
        pan3button2 = pan3button.Clone
        pan3button2.Text = "Anipath"
        pan3button2.LargeImage = Images.getBitmap(My.Resources.large11)
        pan3button2.Tag = "_pline"
        Dim pan3button3 As RibbonButton = New RibbonButton()
        pan3button3 = pan3button.Clone
        pan3button3.Text = "Sun" + vbCrLf + "Options"
        pan3button3.LargeImage = Images.getBitmap(My.Resources.large12)
        pan3button3.Tag = "_pline"
        Dim pan3button4 As RibbonButton = New RibbonButton()
        pan3button4 = pan3button.Clone
        pan3button4.Text = "Render" + vbCrLf + "Options"
        pan3button4.LargeImage = Images.getBitmap(My.Resources.large13)
        pan3button4.Tag = "_pline"
        Dim pan3button5 As RibbonButton = New RibbonButton()
        pan3button5 = pan3button.Clone
        pan3button5.Text = "Render"
        pan3button5.LargeImage = Images.getBitmap(My.Resources.large14)
        pan3button5.Tag = "_pline"
        panel3Panel.Items.Add(pan3button1)
        panel3Panel.Items.Add(pan3button2)
        panel3Panel.Items.Add(New RibbonSeparator())
        panel3Panel.Items.Add(pan3button3)
        panel3Panel.Items.Add(New RibbonSeparator())
        panel3Panel.Items.Add(pan3button4)
        panel3Panel.Items.Add(pan3button5)
        '////////////////////////////////// P A N E L 4 /////////////////////////////////////
        Dim panel4Panel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
        panel4Panel.Title = "About"
        Dim Panel4 As RibbonPanel = New RibbonPanel()
        Panel4.Source = panel4Panel
        Tab.Panels.Add(Panel4)
        Dim pan4button As RibbonButton = New RibbonButton()
        pan4button = pan1button0.Clone
        Dim pan4button1 As RibbonButton = New RibbonButton()
        pan4button1 = pan4button.Clone
        pan4button1.Text = "Help"
        pan4button1.LargeImage = Images.getBitmap(My.Resources.large15)
        pan4button1.Tag = "_pline"
        panel4Panel.Items.Add(pan4button1)
        Tab.IsActive = True
        Tab.IsVisible = True
    End Sub

    Public Class RibbonCommandHandler

        Implements System.Windows.Input.ICommand

        Public Function CanExecute(ByVal parameter As Object) As Boolean Implements System.Windows.Input.ICommand.CanExecute
            Return True
        End Function

        Public Event CanExecuteChanged(ByVal sender As Object, ByVal e As System.EventArgs) Implements System.Windows.Input.ICommand.CanExecuteChanged

        Public Sub Execute(ByVal parameter As Object) Implements System.Windows.Input.ICommand.Execute

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument

            If TypeOf parameter Is RibbonButton Then
                Dim button As RibbonButton = TryCast(parameter, RibbonButton)
                doc.SendStringToExecute(button.Tag + " ", True, False, False)
            End If
        End Sub
    End Class

    Public Class Images

        Public Shared Function getBitmap(ByVal image As Bitmap) As BitmapImage
            Dim stream As New MemoryStream()
            image.Save(stream, ImageFormat.Png)
            Dim bmp As New BitmapImage()
            bmp.BeginInit()
            bmp.StreamSource = stream
            bmp.EndInit()
            Return bmp
        End Function
    End Class
End Class
