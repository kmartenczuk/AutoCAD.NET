'https://www.cadcenter.pl/
'info@cadcenter.pl
'Autodesk Authorized Developer Kamil Martenczuk (ADN ID DEPL2710)
'6/16/2020
'3DStairsGenrator

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Stairs3D

    <CommandMethod("Stairs3D")>
    Public Sub Action()
        Dim ccpDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ccpPtRes1 As PromptPointResult : Dim ccpPtRes2 As PromptPointResult : Dim ccpMenuResult As PromptResult : Dim ccpHeightResult As PromptDoubleResult
        Dim ccpMenu As PromptKeywordOptions = New PromptKeywordOptions("") : Dim ccpPtOpt1 As PromptPointOptions = New PromptPointOptions("")
        Dim ccpPtOpt2 As PromptPointOptions = New PromptPointOptions("") : Dim ccpQStOpt As PromptDoubleOptions = New PromptDoubleOptions("")
        Dim ccpHeight As PromptDoubleOptions = New PromptDoubleOptions("")

        '*********************************************************************
        '------------------------------- M e n u -----------------------------
        '*********************************************************************
        ccpMenu.Message = vbCrLf + "Select Style of stairs: "
        ccpMenu.Keywords.Add("Classic")
        ccpMenu.Keywords.Add("Light")
        ccpMenu.Keywords.Default = "Classic"
        ccpMenu.AllowNone = True
        ccpMenuResult = ccpDoc.Editor.GetKeywords(ccpMenu)
        If ccpMenuResult.Status = PromptStatus.Cancel Then Exit Sub

        '*********************************************************************
        '--------------------- S e l e c t   P o i n t s ---------------------
        '*********************************************************************
        ccpPtOpt1.Message = vbCrLf + "Start piont (down-left corner: )"
        ccpPtRes1 = ccpDoc.Editor.GetPoint(ccpPtOpt1)
        ccpPtOpt2.Message = vbCrLf + "End point (up-right corner: )"
        ccpPtRes2 = ccpDoc.Editor.GetPoint(ccpPtOpt2)
        If ccpPtRes2.Value.X = ccpPtRes1.Value.X Or ccpPtRes2.Value.Y = ccpPtRes1.Value.Y Or ccpPtRes2.Value.Z < ccpPtRes1.Value.Z Then
            MsgBox("Wrong Width or Height" + vbCrLf + "Value must be > 0" + vbCrLf + "Please try again", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        '*********************************************************************
        '------------ E n t e r  q u a n t i t y   o f   S t e p s -----------
        '*********************************************************************
        ccpQStOpt.Message = vbCrLf + "Quantity of Steps: "
        Dim ccpQStRes As PromptDoubleResult = ccpDoc.Editor.GetDouble(ccpQStOpt)
        Dim ccpQSt As Double = ccpQStRes.Value
        Dim ccpL As Double = 0 : Dim ccpW As Double = 0
        If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
            ccpL = (ccpPtRes2.Value.Y - ccpPtRes1.Value.Y) / ccpQSt
            ccpW = ccpPtRes2.Value.X - ccpPtRes1.Value.X
        Else
            ccpL = (ccpPtRes2.Value.X - ccpPtRes1.Value.X) / ccpQSt
            ccpW = ccpPtRes2.Value.Y - ccpPtRes1.Value.Y
        End If

        '*********************************************************************
        '-------------------- E n t e r   H e i g h t   i f ------------------
        '*********************************************************************
        Dim ccpH As Double
        If ccpPtRes2.Value.Z = ccpPtRes1.Value.Z Then
            ccpHeight.Message = "Enter height: "
            ccpHeightResult = ccpDoc.Editor.GetDouble(ccpHeight)
            ccpH = ccpHeightResult.Value
            If ccpHeightResult.Status = PromptStatus.Cancel Then Exit Sub
        ElseIf ccpPtRes2.Value.Z > ccpPtRes1.Value.Z Then
            ccpH = ccpPtRes2.Value.Z - ccpPtRes1.Value.Z
        Else
            MsgBox("Wrong value of Height" + vbCrLf + "It must be >= 0" + vbCrLf + "Please try again", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If ccpQStRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim ccpHSt As Double = ccpH / ccpQSt
        Dim ccpDB As Database = ccpDoc.Database

        '*********************************************************************
        '----------------- S t a i r s    G e n e r a t o r ------------------
        '*********************************************************************

        If ccpMenuResult.StringResult = "Classic" Then
            '*********************************************************************
            '-------------------- C l a s s i c   S t a i r s  -------------------
            '*********************************************************************
            Using ccpTrans As Transaction = ccpDB.TransactionManager.StartTransaction
                Dim ccpBlk As BlockTable
                ccpBlk = ccpTrans.GetObject(ccpDB.BlockTableId, OpenMode.ForRead)
                Dim ccpBlkRec As BlockTableRecord
                ccpBlkRec = ccpTrans.GetObject(ccpBlk(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim ccpPolyAlfa As Polyline3d = New Polyline3d()
                ccpPolyAlfa.SetDatabaseDefaults()
                ccpPolyAlfa.ColorIndex = 9
                ccpBlkRec.AppendEntity(ccpPolyAlfa)
                Dim ccpPolyAlfaPts As Point3dCollection = New Point3dCollection()
                Dim ccpCount As Integer = 0
                If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> Y-axis long option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL, 0))
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y, 0))
                    While (ccpCount < ccpQSt)
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + (ccpL * ccpCount), ccpHSt * ccpCount))
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + (ccpL * ccpCount), ccpHSt * (ccpCount + 1)))
                        If ccpCount < ccpQSt - 1 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL * (ccpCount + 1), ccpHSt * (ccpCount + 1)))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL * (ccpCount + 1), ccpHSt * (ccpCount + 1)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL * (ccpCount + 1), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL * (ccpCount + 2), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                        End If
                        ccpCount = ccpCount + 1
                    End While
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL, 0))
                Else
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> X-axis long option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y, 0))
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y, 0))
                    While (ccpCount < ccpQSt)
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * ccpCount), ccpPtRes1.Value.Y, ccpHSt * ccpCount))
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * ccpCount), ccpPtRes1.Value.Y, ccpHSt * (ccpCount + 1)))
                        If ccpCount < ccpQSt - 1 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1), ccpPtRes1.Value.Y, ccpHSt * (ccpCount + 1)))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1), ccpPtRes1.Value.Y, ccpHSt * (ccpCount + 1)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1), ccpPtRes1.Value.Y, ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 2), ccpPtRes1.Value.Y, ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                        End If
                        ccpCount = ccpCount + 1
                    End While
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y, 0))
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '>>>>>>>>>> A p p e n d >>>>>>>>>
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                For Each ccpPt3d As Point3d In ccpPolyAlfaPts
                    Dim ccpPolyVer3d As PolylineVertex3d = New PolylineVertex3d(ccpPt3d)
                    ccpPolyAlfa.AppendVertex(ccpPolyVer3d)
                    ccpTrans.AddNewlyCreatedDBObject(ccpPolyVer3d, True)
                Next
                ccpPolyAlfa.TransformBy(Matrix3d.Displacement(New Vector3d(0, 0, ccpPtRes1.Value.Z)))
                Dim ccp3DAlfa As Solid3d = New Solid3d()
                If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                    ccp3DAlfa.CreateExtrudedSolid(ccpPolyAlfa, New Vector3d(ccpW.ToString, 0, 0), New SweepOptions())
                Else
                    ccp3DAlfa.CreateExtrudedSolid(ccpPolyAlfa, New Vector3d(0, ccpW.ToString, 0), New SweepOptions())
                End If
                ccp3DAlfa.ColorIndex = 9
                ccpBlkRec.AppendEntity(ccp3DAlfa)
                ccpTrans.AddNewlyCreatedDBObject(ccp3DAlfa, True)
                ccpPolyAlfa.Erase(True)
                ccpTrans.Commit()
            End Using
            '*********************************************************************
            '---------------- E n d   C l a s s i c   S t a i r s  ---------------
            '*********************************************************************

        ElseIf ccpMenuResult.StringResult = "Light" Then
            '*********************************************************************
            '----------------------- L i g h t    S t a i r s --------------------
            '*********************************************************************
            Using ccpTrans As Transaction = ccpDB.TransactionManager.StartTransaction
                Dim ccpBlk As BlockTable
                ccpBlk = ccpTrans.GetObject(ccpDB.BlockTableId, OpenMode.ForRead)
                Dim ccpBlkRec As BlockTableRecord
                ccpBlkRec = ccpTrans.GetObject(ccpBlk(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '>>>>>>>>>>>> S t e p s >>>>>>>>>>
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                Dim ccpPoly As Polyline = New Polyline()
                ccpPoly.SetDatabaseDefaults()
                If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> Y-axis long option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPoly.AddVertexAt(0, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y), 0, 0, 0)
                    ccpPoly.AddVertexAt(1, New Point2d(ccpPtRes1.Value.X + ccpW, ccpPtRes1.Value.Y), 0, 0, 0)
                    ccpPoly.AddVertexAt(2, New Point2d(ccpPtRes1.Value.X + ccpW, ccpPtRes1.Value.Y + ccpL), 0, 0, 0)
                    ccpPoly.AddVertexAt(3, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpL), 0, 0, 0)
                    ccpPoly.AddVertexAt(4, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y), 0, 0, 0)
                Else
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> X-axis long option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPoly.AddVertexAt(0, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y), 0, 0, 0)
                    ccpPoly.AddVertexAt(1, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y + ccpW), 0, 0, 0)
                    ccpPoly.AddVertexAt(2, New Point2d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y + ccpW), 0, 0, 0)
                    ccpPoly.AddVertexAt(3, New Point2d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y), 0, 0, 0)
                    ccpPoly.AddVertexAt(4, New Point2d(ccpPtRes1.Value.X, ccpPtRes1.Value.Y), 0, 0, 0)
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '>>>>>>>>>> C r e a t e   &   A p p e n d >>>>>>>>>
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                ccpPoly.TransformBy(Matrix3d.Displacement(New Vector3d(0, 0, ccpPtRes1.Value.Z + (3 * ccpHSt.ToString / 4))))
                Dim ccp3D As Solid3d = New Solid3d()
                ccp3D.CreateExtrudedSolid(ccpPoly, New Vector3d(0, 0, ccpHSt.ToString / 4), New SweepOptions())
                ccp3D.ColorIndex = 16
                Dim ccpColl As DBObjectCollection = New DBObjectCollection()
                ccpColl.Add(ccp3D)
                Dim nCount As Integer = 0
                While (nCount < ccpQSt)
                    Dim ccp3DClone As Solid3d = ccp3D.Clone()
                    If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                        ccp3DClone.TransformBy(Matrix3d.Displacement(New Vector3d(0, Str(nCount * ccpL), (nCount * ccpHSt))))
                        ccpColl.Add(ccp3DClone)
                    Else
                        ccp3DClone.TransformBy(Matrix3d.Displacement(New Vector3d(Str(nCount * ccpL), 0, (nCount * ccpHSt))))
                        ccpColl.Add(ccp3DClone)
                    End If
                    ccpBlkRec.AppendEntity(ccp3DClone)
                    ccpTrans.AddNewlyCreatedDBObject(ccp3DClone, True)
                    nCount = nCount + 1
                End While
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '>>>>>>>>>> C e n t e r >>>>>>>>>
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                Dim ccpPolyAlfa As Polyline3d = New Polyline3d()
                ccpPolyAlfa.SetDatabaseDefaults()
                ccpBlkRec.AppendEntity(ccpPolyAlfa)
                Dim ccpPolyAlfaPts As Point3dCollection = New Point3dCollection()
                Dim ccpCount As Integer = 0
                If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> Y-axis long center elem. option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL, 0 + ccpHSt.ToString / 4))
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + (ccpL * 0.1), 0 + ccpHSt.ToString / 4))
                    While (ccpCount < ccpQSt)
                        If ccpCount = 0 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + (ccpL * ccpCount) + (ccpL * 0.1), ccpHSt * ccpCount + ccpHSt.ToString / 4))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + (ccpL * ccpCount) + (ccpL * 0.1), ccpHSt * ccpCount))
                        End If
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + (ccpL * ccpCount) + (ccpL * 0.1), ccpHSt * (ccpCount + 1)))
                        If ccpCount < ccpQSt - 1 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL * (ccpCount + 1) + (ccpL * 0.1), ccpHSt * (ccpCount + 1)))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL * (ccpCount + 1), ccpHSt * (ccpCount + 1)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL * (ccpCount + 1), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL * (ccpCount + 2) + (ccpL * 0.1), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                        End If
                        ccpCount = ccpCount + 1
                    End While
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (3 * ccpW / 8), ccpPtRes1.Value.Y + ccpL, 0 + ccpHSt.ToString / 4))
                Else
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    '>>>>>>>>>>>> X-axis long center elem. option >>>>>>>>>
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y + (3 * ccpW / 8), 0 + ccpHSt.ToString / 4))
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), 0 + ccpHSt.ToString / 4))
                    While (ccpCount < ccpQSt)
                        If ccpCount = 0 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * ccpCount) + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * ccpCount + ccpHSt.ToString / 4))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * ccpCount) + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * ccpCount))
                        End If
                        ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + (ccpL * ccpCount) + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * (ccpCount + 1)))
                        If ccpCount < ccpQSt - 1 Then
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1) + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * (ccpCount + 1)))
                        Else
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * (ccpCount + 1)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                            ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL * (ccpCount + 2) + (ccpL * 0.1), ccpPtRes1.Value.Y + (3 * ccpW / 8), ccpHSt * (ccpCount + 1) - (0.1 * ccpHSt)))
                        End If
                        ccpCount = ccpCount + 1
                    End While
                    ccpPolyAlfaPts.Add(New Point3d(ccpPtRes1.Value.X + ccpL, ccpPtRes1.Value.Y + (3 * ccpW / 8), 0 + ccpHSt.ToString / 4))
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '>>>>>>>>>> C r e a t e   &   A p p e n d >>>>>>>>>
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                For Each ccpPt3d As Point3d In ccpPolyAlfaPts
                    Dim ccpPolyVer3d As PolylineVertex3d = New PolylineVertex3d(ccpPt3d)
                    ccpPolyAlfa.AppendVertex(ccpPolyVer3d)
                    ccpTrans.AddNewlyCreatedDBObject(ccpPolyVer3d, True)
                Next
                ccpPolyAlfa.TransformBy(Matrix3d.Displacement(New Vector3d(0, 0, ccpPtRes1.Value.Z - ccpHSt.ToString / 4)))
                Dim ccp3DAlfa As Solid3d = New Solid3d()
                ccp3DAlfa.ColorIndex = 9
                If (ccpPtRes2.Value.X > ccpPtRes1.Value.X And ccpPtRes2.Value.Y > ccpPtRes1.Value.Y) Or (ccpPtRes2.Value.X < ccpPtRes1.Value.X And ccpPtRes2.Value.Y < ccpPtRes1.Value.Y) Then
                    ccp3DAlfa.CreateExtrudedSolid(ccpPolyAlfa, New Vector3d(ccpW.ToString / 4, 0, 0), New SweepOptions())
                Else
                    ccp3DAlfa.CreateExtrudedSolid(ccpPolyAlfa, New Vector3d(0, ccpW.ToString / 4, 0), New SweepOptions())
                End If
                ccpBlkRec.AppendEntity(ccp3DAlfa)
                ccpTrans.AddNewlyCreatedDBObject(ccp3DAlfa, True)
                ccpPolyAlfa.Erase(True)
                ccpTrans.Commit()
            End Using
            '*********************************************************************
            '------------------ E n d    L i g h t    S t a i r s ----------------
            '*********************************************************************
        End If
    End Sub
End Class