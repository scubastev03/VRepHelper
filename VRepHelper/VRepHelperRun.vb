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
        Dim oAddRepIfNotExist As Boolean = False
        Dim oCmpOccCount As Double = oCmpOccs.AllLeafOccurrences().Count
        Dim oMinCount As Double = 2500
        Try
            Select Case MessageBox.Show("Continue without adding non existing View and LOD Representations?", oCaption, MessageBoxButtons.YesNo)
                Case DialogResult.No
                    oAddRepIfNotExist = True
                Case Else
            End Select
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
                    oVRepTable = oDataSet.Tables("ViewRep") 'Get View Rep table from referenced xml db file
                    oVRepTable.PrimaryKey = New DataColumn() {oVRepTable.Columns("Name")}
                    oCDefTable = oDataSet.Tables("ComponentDef")
                    oCDefTable.PrimaryKey = New DataColumn() {oCDefTable.Columns("Name")}
                    For Each oVRepRow As DataRow In oVRepTable.Rows 'Create View Representations If Not Exist
                        If oVRepRow.Item("DVR") Then
                            Try
                                oViewRep = oViewReps.Item(oVRepRow.Item("Name"))
                            Catch
                                If oAddRepIfNotExist Then
                                    oViewRep = oViewReps.Add(oVRepRow.Item("Name"))
                                End If
                            End Try
                        End If
                    Next oVRepRow
                    For Each oViewRep In oViewReps 'Run Through Each View Representation and Set Visibility
                        If oViewRep.DesignViewType() = DesignViewTypeEnum.kPublicDesignViewType Then
                            Dim oViewRepId As String
                            Try
                                oViewRepId = oVRepTable.Rows.Find(oViewRep.Name)(1).ToString()
                            Catch
                                Continue For
                            End Try
                            oViewRep.Activate() 'Activate View Representation
                            oInvApp.ScreenUpdating = True
                            oInvApp.StatusBarText() = "Setting " & oViewRep.Name & " to Active Design View Representation"
                            oInvApp.ScreenUpdating = False
                            If Not oAsmCmpDef.RepresentationsManager.ActiveLevelOfDetailRepresentation().Name = "Master" Then
                                oLODReps.Item("Master").Activate()
                            End If
                            For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences 'Get Leaf Occurences for Current Assembly Document Definition
                                Select Case oCmpOcc.DefinitionDocumentType()
                                    Case DocumentTypeEnum.kAssemblyDocumentObject
                                        If TypeOf oCmpOcc.Definition Is VirtualComponentDefinition Then
                                            Continue For
                                        End If
                                        If oCmpOcc.BOMStructure() = BOMStructureEnum.kPhantomBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                            If oMinCount >= oCmpOccCount AndAlso Not oMinCount = 0 Then
                                                If oCmpOcc.IsAssociativeToDesignViewRepresentation() Then
                                                    If Not oCmpOcc.ActiveDesignViewRepresentation() = oViewRep.Name Then
                                                        Try
                                                            oCmpOcc.SetDesignViewRepresentation(oViewRep.Name, True)
                                                        Catch
                                                            Continue For
                                                        End Try
                                                    End If
                                                Else
                                                    Try
                                                        oCmpOcc.SetDesignViewRepresentation(oViewRep.Name, True)
                                                    Catch
                                                        Continue For
                                                    End Try
                                                End If
                                            Else
                                                Try
                                                    oCmpOcc.SetDesignViewRepresentation(oViewRep.Name, False)
                                                Catch
                                                    Continue For
                                                End Try
                                            End If
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
                                                                If Not oCmpOcc.Visible Then
                                                                    oObjCol1.Add(oCmpOcc)
                                                                End If
                                                            ElseIf Not oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oViewRepId) Then
                                                                If oCmpOcc.Visible Then
                                                                    oObjCol2.Add(oCmpOcc)
                                                                End If
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
                            If oObjCol1.Count() Then
                                Try
                                    oViewRep.SetVisibilityOfOccurrences(oObjCol1, True)
                                    oInvApp.ScreenUpdating = True
                                    oInvApp.StatusBarText() = "Setting Component Occurrence Collection Visibility to True"
                                    oInvApp.ScreenUpdating = False
                                Catch ex As Exception
                                    MsgBox(ex.ToString,, oCaption)
                                Finally
                                    oObjCol1.Clear()
                                End Try
                            End If
                            If oObjCol2.Count() Then
                                Try
                                    oViewRep.SetVisibilityOfOccurrences(oObjCol2, False)
                                    oInvApp.ScreenUpdating = True
                                    oInvApp.StatusBarText() = "Setting Component Occurrence Collection Visibility to False"
                                    oInvApp.ScreenUpdating = False
                                Catch ex As Exception
                                    MsgBox(ex.ToString,, oCaption)
                                Finally
                                    oObjCol2.Clear()
                                End Try
                            End If
                        End If
                    Next oViewRep
                    For Each oVRepRow As DataRow In oVRepTable.Rows 'Create View Representations If Not Exist
                        If oVRepRow.Item("LOD") Then
                            Try
                                oLODRep = oLODReps.Item(oVRepRow.Item("Name"))
                            Catch
                                If oAddRepIfNotExist Then
                                    oLODRep = oLODReps.Add(oVRepRow.Item("Name"))
                                End If
                            End Try
                        End If
                    Next oVRepRow
                    For Each oLODRep In oLODReps 'Run Through Each View Representation and Set Visibility
                        If oLODRep.LevelOfDetail() = LevelOfDetailEnum.kCustomLevelOfDetail Then
                            Dim oLODRepId As String
                            Try
                                oLODRepId = oVRepTable.Rows.Find(oLODRep.Name)(1).ToString()
                            Catch
                                Continue For
                            End Try
                            oLODRep.Activate() 'Activate View Representation
                            oInvApp.ScreenUpdating = True
                            oInvApp.StatusBarText() = "Setting " & oLODRep.Name & " to Active Level of Detail Representation"
                            oInvApp.ScreenUpdating = False
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
                                                                If oCmpOcc.Suppressed Then
                                                                    oCmpOcc.Unsuppress()
                                                                End If
                                                            ElseIf Not oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oLODRepId) Then
                                                                If Not oCmpOcc.Suppressed Then
                                                                    oCmpOcc.Suppress()
                                                                End If
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
                            oDoc.Save2(False)
                        End If
                    Next oLODRep
                    oInvApp.ScreenUpdating = True
                    oInvApp.StatusBarText() = "Complete"
                    oInvApp.ScreenUpdating = False
                End If
                oStopWatch.Stop()
                oDoc.Save2(False)
                MsgBox("Setting Design View & LOD Representation Process Completed." & vbCr & "Process Completed in " & oStopWatch.Elapsed.TotalMinutes.ToString("F2") & " Minutes.",, oCaption)
            Catch ex As Exception
                MsgBox(ex.ToString,, oCaption)
            End Try
            Select Case MessageBox.Show("Do you want to set View Representation and Level of Detail to the original values?", oCaption, MessageBoxButtons.YesNo)
                Case DialogResult.Yes
                    oInvApp.ScreenUpdating = True
                    oInvApp.StatusBarText() = "Setting View Representation and Level of Detail Representations to initial values."
                    oInvApp.ScreenUpdating = False
                    oViewRepActive.Activate()
                    oLODRepActive.Activate()
                Case Else
            End Select
        Catch
            MsgBox("Sorry! Something did not go right.",, oCaption)
        Finally
            oInvApp.ScreenUpdating = True
            oInvApp.StatusBarText() = "Complete"
            oInvApp.AssemblyOptions.DeferUpdate = oDeferUpdate
        End Try
    End Sub
End Class