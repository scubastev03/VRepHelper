Imports Inventor
Imports System.Data
Public Class VRepHelperRun
    Public Sub Run()
        Dim oDataSet As New DataSet()
        Dim oVRepTable As New DataTable()
        Dim oCDefTable As New DataTable()
        Dim oTimerStart As DateTime = DateAndTime.Now()
        Dim oInvApp As Application = VRepHelper.StandardAddInServer.oInvApp
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
        Dim oFileCount As Integer = 0
        Try
            oDataSet.ReadXml("VRepHelperDb.xml", XmlReadMode.ReadSchema)
        Catch
            MsgBox("Sorry, no VRepHelperDb.xml file is present.",, "Design View & LOD Representations")
            Exit Sub
        End Try
        Try
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then 'Check If File Is An Assembly
                oFileCount = oDoc.AllReferencedDocuments.Count() 'Get count of referenced files
                'MsgBox(oFileCount.tostring)
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
                        oObjCol1.Clear() 'Clear Object Collections
                        oObjCol2.Clear() 'Clear Object Collections
                        For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences 'Get Leaf Occurences for Current Assembly Document Definition
                            If oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kAssemblyDocumentObject Then
                                '					Continue For
                                If oCmpOcc.BOMStructure() = BOMStructureEnum.kPhantomBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                    Try
                                        If oCmpOcc.ActiveDesignViewRepresentation() = oViewRep.Name Then
                                            Continue For
                                        ElseIf oCmpOcc.ActiveDesignViewRepresentation() <> oViewRep.Name Then
                                            Try
                                                'oCmpOcc.IsAssociativeToDesignViewRepresentation() = False
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
                            ElseIf oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kPartDocumentObject Then
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
                                                        'oCmpOcc.Visible = True
                                                        oObjCol1.Add(oCmpOcc)
                                                    ElseIf Not oCDefTable.Rows.Find(oPropSets.Item("Inventor Summary Information").Item("Title").Value)(oViewRepId) Then
                                                        'oCmpOcc.Visible = False
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
                            Else
                                Continue For
                            End If
                        Next oCmpOcc
                        If oObjCol1.Count() <> 0 Then
                            Try
                                oViewRep.SetVisibilityOfOccurrences(oObjCol1, True)
                            Catch ex As Exception
                                MsgBox(ex.ToString,, "Design View & LOD Representations")
                            End Try
                        End If
                        If oObjCol2.Count() <> 0 Then
                            Try
                                oViewRep.SetVisibilityOfOccurrences(oObjCol2, False)
                            Catch ex As Exception
                                MsgBox(ex.ToString,, "Design View & LOD Representations")
                            End Try
                        End If
                        oInvApp.StatusBarText() = oViewRep.Name & " Complete... Moving to Next Design View Representation."
                        ' Try
                        ' oLODReps.Item(oViewRep.Name).Delete()
                        ' Catch
                        ' End Try
                        ' If oFileCount > 50 Then
                        ' Try
                        ' oViewReps.Item(oViewRep.Name).CopyToLevelOfDetail(oViewRep.Name)
                        ' Catch ex As Exception
                        ' 'MsgBox(ex.tostring)
                        ' End Try
                        ' End If
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
                        oInvApp.StatusBarText() = "Setting " & oLODRep.Name & " to active Level of Detail Representation"
                        For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences 'Get Leaf Occurences for Current Assembly Document Definition
                            If oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kAssemblyDocumentObject Then
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
                            ElseIf oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kPartDocumentObject Then
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
                            Else
                                Continue For
                            End If
                        Next oCmpOcc
                        oDoc.Save()
                        oInvApp.StatusBarText() = oLODRep.Name & " Complete... Moving to Next Level of Detail Representation."
                        ' Try
                        ' oLODReps.Item(oViewRep.Name).Delete()
                        ' Catch
                        ' End Try
                        ' If oFileCount > 50 Then
                        ' Try
                        ' oViewReps.Item(oViewRep.Name).CopyToLevelOfDetail(oViewRep.Name)
                        ' Catch ex As Exception
                        ' 'MsgBox(ex.tostring)
                        ' End Try
                        ' End If
                    End If
                Next oLODRep
            End If

            oLODRepActive.Activate()
            Dim TimeSpent As System.TimeSpan = Now.Subtract(oTimerStart)
            oDoc.Save()
            MsgBox("Setting Design View & LOD Representation Process Completed." & vbCr & "Process Completed in " & TimeSpent.Seconds & " Seconds.",, "Design View & LOD Representations")
        Catch ex As Exception
            MsgBox(ex.ToString,, "Design View & LOD Representations")
        End Try
        oDoc = Nothing
    End Sub

End Class
