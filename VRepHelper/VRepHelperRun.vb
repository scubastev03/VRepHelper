Imports Inventor
Imports System.Data
Imports System.Windows.Forms
Public Class VRepHelperRun
    Public Sub Run()
        Dim oDataSet As New DataSet()
        Const oCaption As String = "Design View & LOD Representations"
        Dim oVRepTable As New DataTable()
        Dim oCDefTable As New DataTable()
        Dim oStopWatch As New Stopwatch()
        Dim oInvApp As Inventor.Application = VRepHelper.StandardAddInServer.oInvApp
        Dim oDoc As Document = oInvApp.ActiveDocument
        Dim oAsmCmpDef As AssemblyComponentDefinition = oDoc.ComponentDefinition
        Dim oCmpOccs As ComponentOccurrences = oAsmCmpDef.Occurrences
        Dim oViewReps As DesignViewRepresentations = oAsmCmpDef.RepresentationsManager.DesignViewRepresentations
        Dim oViewRep As DesignViewRepresentation
        Dim oViewRepActive As DesignViewRepresentation = oAsmCmpDef.RepresentationsManager.ActiveDesignViewRepresentation()
        Dim oLODReps As LevelOfDetailRepresentations = oAsmCmpDef.RepresentationsManager.LevelOfDetailRepresentations
        Dim oLODRep As LevelOfDetailRepresentation
        Dim oLODRepActive As LevelOfDetailRepresentation = oAsmCmpDef.RepresentationsManager.ActiveLevelOfDetailRepresentation()
        Dim oPropSets As PropertySets
        Dim oObjCol1 As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection
        Dim oObjCol2 As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection
        Dim oDeferUpdate As Boolean = oInvApp.AssemblyOptions.DeferUpdate
        Dim oFileCount As Integer = 0
        Try
            Select Case MessageBox.Show("Do you want to proceed? This process may take several minutes.", oCaption, MessageBoxButtons.OKCancel)
                Case DialogResult.Cancel
                    Exit Sub
                Case Else
            End Select
            oStopWatch.Start()
            oInvApp.ScreenUpdating = False
            oInvApp.AssemblyOptions.DeferUpdate = True
            Try
                oDataSet.ReadXml(oInvApp.DesignProjectManager.ResolveFile("C:\", "VRepHelperDb.xml"), XmlReadMode.ReadSchema)
            Catch ex As ArgumentNullException
                oDataSet.ReadXml(My.Application.Info.DirectoryPath & "\VRepHelperDb.xml", XmlReadMode.ReadSchema)
            Catch
                MsgBox("Sorry, no VRepHelperDb.xml file is present.",, oCaption)
                Exit Sub
            End Try
            Try
                If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then 'Check If File Is An Assembly
                    oFileCount = oDoc.AllReferencedDocuments.Count() 'Get count of referenced files
                    oVRepTable = oDataSet.Tables("ViewRep") 'Get View Rep table from referenced xml db file
                    oVRepTable.PrimaryKey = New DataColumn() {oVRepTable.Columns("Name")}
                    oCDefTable = oDataSet.Tables("ComponentDef")
                    oCDefTable.PrimaryKey = New DataColumn() {oCDefTable.Columns("Name")}
                    For Each oVRepRow As DataRow In oVRepTable.Rows 'Create View Representations If Not Exist
                        If oVRepRow.Item("DVR") Then
                            Try
                                oViewReps.Item(oVRepRow.Item("Name")).Activate()
                            Catch
                                oViewRep = oViewReps.Add(oVRepRow.Item("Name"))
                            End Try
                        End If
                    Next oVRepRow
                    For Each oViewRep In oViewReps 'Run Through Each View Representation and Set Visibility
                        If oViewRep.DesignViewType() = DesignViewTypeEnum.kPublicDesignViewType Then
                            oViewRep.Activate() 'Activate View Representation
                            Dim oViewRepId As String
                            Try
                                oViewRepId = oVRepTable.Rows.Find(oViewRep.Name)(1).ToString()
                            Catch
                                Continue For
                            End Try
                            oInvApp.StatusBarText() = "Setting " & oViewRep.Name & " to Active Design View Representation"
                            oLODReps.Item("Master").Activate()
                            For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences 'Get Leaf Occurences for Current Assembly Document Definition
                                Select Case oCmpOcc.DefinitionDocumentType()
                                    Case DocumentTypeEnum.kAssemblyDocumentObject
                                        If TypeOf oCmpOcc.Definition Is VirtualComponentDefinition Then
                                            Continue For
                                        End If
                                        If oCmpOcc.BOMStructure() = BOMStructureEnum.kPhantomBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                            Try
                                                If oCmpOcc.ActiveDesignViewRepresentation() = oViewRep.Name Then
                                                    Continue For
                                                ElseIf oCmpOcc.ActiveDesignViewRepresentation() <> oViewRep.Name Then
                                                    Try
                                                        oCmpOcc.SetDesignViewRepresentation(oViewRep.Name)
                                                    Catch
                                                        Continue For
                                                    End Try
                                                ElseIf String.IsNullOrEmpty(oCmpOcc.ActiveDesignViewRepresentation()) Then
                                                    Try
                                                        oCmpOcc.SetDesignViewRepresentation(oViewRep.Name)
                                                    Catch
                                                        Continue For
                                                    End Try
                                                End If
                                            Catch
                                                Continue For
                                            End Try
                                        Else
                                            Continue For
                                        End If
                                    Case DocumentTypeEnum.kPartDocumentObject
                                        If oCmpOcc.BOMStructure() = BOMStructureEnum.kNormalBOMStructure Or BOMStructureEnum.kPurchasedBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                            Try
                                                oPropSets = oCmpOcc.Definition.Document.PropertySets
                                                If Not oPropSets.PropertySetExists("Inventor Summary Information") Then
                                                    Continue For
                                                Else
                                                    Try
                                                        If String.IsNullOrEmpty(oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oViewRepId)) Then
                                                            Continue For
                                                        Else
                                                            If oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oViewRepId) Then
                                                                oObjCol1.Add(oCmpOcc)
                                                            ElseIf Not oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oViewRepId) Then
                                                                oObjCol2.Add(oCmpOcc)
                                                            End If
                                                        End If
                                                    Catch
                                                        Continue For
                                                    End Try
                                                End If
                                            Catch
                                            End Try
                                        End If
                                    Case Else
                                        Continue For
                                End Select
                            Next oCmpOcc
                            If oObjCol1.Count() <> 0 Then
                                Try
                                    oViewRep.SetVisibilityOfOccurrences(oObjCol1, True)
                                Catch ex As Exception
                                    MsgBox(ex.ToString,, oCaption)
                                Finally
                                    oObjCol1.Clear()
                                End Try
                            End If
                            If oObjCol2.Count() <> 0 Then
                                Try
                                    oViewRep.SetVisibilityOfOccurrences(oObjCol2, False)
                                Catch ex As Exception
                                    MsgBox(ex.ToString,, oCaption)
                                Finally
                                    oObjCol2.Clear()
                                End Try
                            End If
                            oViewRepActive.Activate()
                        End If
                        For Each oVRepRow As DataRow In oVRepTable.Rows 'Create View Representations If Not Exist
                            If oVRepRow.Item("LOD") Then
                                Try
                                    oLODReps.Item(oVRepRow.Item("Name")).Activate()
                                Catch
                                    oLODRep = oLODReps.Add(oVRepRow.Item("Name"))
                                End Try
                            End If
                        Next oVRepRow
                    Next oViewRep
                    For Each oLODRep In oLODReps 'Run Through Each View Representation and Set Visibility
                        If oLODRep.LevelOfDetail() = LevelOfDetailEnum.kCustomLevelOfDetail Then
                            oLODRep.Activate() 'Activate View Representation
                            Dim oLODRepId As String
                            Try
                                oLODRepId = oVRepTable.Rows.Find(oLODRep.Name)(1).ToString()
                            Catch
                                Continue For
                            End Try
                            For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences 'Get Leaf Occurences for Current Assembly Document Definition
                                Select Case oCmpOcc.DefinitionDocumentType()
                                    Case DocumentTypeEnum.kAssemblyDocumentObject
                                        If TypeOf oCmpOcc.Definition Is VirtualComponentDefinition Then
                                            Continue For
                                        End If
                                        If oCmpOcc.BOMStructure() = BOMStructureEnum.kPhantomBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                            Try
                                                If oCmpOcc.ActiveLevelOfDetailRepresentation() = oLODRep.Name Then
                                                    Continue For
                                                Else
                                                    Try
                                                        oCmpOcc.SetLevelOfDetailRepresentation(oLODRep.Name)
                                                    Catch
                                                        Continue For
                                                    End Try
                                                End If
                                            Catch
                                                Continue For
                                            End Try
                                        Else
                                            Continue For
                                        End If
                                    Case DocumentTypeEnum.kPartDocumentObject
                                        If oCmpOcc.BOMStructure() = BOMStructureEnum.kNormalBOMStructure Or BOMStructureEnum.kPurchasedBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Or BOMStructureEnum.kPhantomBOMStructure Then
                                            Try
                                                oPropSets = oCmpOcc.Definition.Document.PropertySets
                                                If Not oPropSets.PropertySetExists("Inventor Summary Information") Then
                                                    Continue For
                                                Else
                                                    Try
                                                        If String.IsNullOrEmpty(oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oLODRepId)) Then
                                                            Continue For
                                                        Else
                                                            If oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oLODRepId) Then
                                                                oCmpOcc.Unsuppress()
                                                            ElseIf Not oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oLODRepId) Then
                                                                oCmpOcc.Suppress()
                                                            End If
                                                        End If
                                                    Catch
                                                        Continue For
                                                    End Try
                                                End If
                                            Catch
                                            End Try
                                        End If
                                    Case Else
                                        Continue For
                                End Select
                            Next oCmpOcc
                            oDoc.Save()
                        End If
                    Next oLODRep
                End If
                oLODRepActive.Activate()
                oStopWatch.Stop()
                oDoc.Save()
                MsgBox("Setting Design View & LOD Representation Process Completed." & vbCr & "Process Completed in " & oStopWatch.Elapsed.TotalMinutes.ToString("F2") & " Minutes.",, oCaption)
            Catch ex As Exception
                MsgBox(ex.ToString,, oCaption)
            End Try
            oDoc = Nothing
        Catch
            MsgBox("Sorry! Something did not go right.",, oCaption)
        Finally
            oInvApp.ScreenUpdating = True
            oInvApp.AssemblyOptions.DeferUpdate = oDeferUpdate
        End Try
    End Sub
End Class