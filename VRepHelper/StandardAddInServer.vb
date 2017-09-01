Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace VRepHelper
    <ProgIdAttribute("VRepHelper.StandardAddInServer"),
    GuidAttribute("730dc371-7758-43e6-82da-85326b8a1911"),
    ComVisible(True)>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application object.
        Public Shared oInvApp As Inventor.Application
        Public Shared o1stTime As Boolean
        Public Shared oInvCLSID As GuidAttribute

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            oInvApp = addInSiteObject.Application
            o1stTime = firstTime

            oInvCLSID = CType(System.Attribute.GetCustomAttribute(GetType(StandardAddInServer),
                                                                     GetType(GuidAttribute)), GuidAttribute)

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.

            Dim oCmdMgr As New CommandManager()

        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            Marshal.ReleaseComObject(oInvApp)
            oInvApp = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()

        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

#End Region

    End Class

    Public Class CommandManager
#Region "Class Variables"
        Private oInvApp As Inventor.Application
        Private oInvCLSID As GuidAttribute
        Private oInvControlDefs As ControlDefinitions
        Private o1stTime As Boolean

#End Region

#Region "Button Definitions" ''Place new button definitions here
        Private WithEvents oVRepRunBtn As ButtonDefinition

#End Region

        Public Sub New()
            Me.oInvApp = VRepHelper.StandardAddInServer.oInvApp
            Me.oInvCLSID = VRepHelper.StandardAddInServer.oInvCLSID
            Me.o1stTime = VRepHelper.StandardAddInServer.o1stTime
            Me.oInvControlDefs = Me.oInvApp.CommandManager.ControlDefinitions
            Me.CreateCommands()
        End Sub

#Region "Create Button Commands" ''Create button commands here
        Public Sub CreateCommands()
            ' Create HDRBIM Project button
            Me.oVRepRunBtn = Me.oInvControlDefs.AddButtonDefinition(
                  "Run VRep Helper",
                  "VRepHelper:CustomCommand:VRepHelperRun",
                  CommandTypesEnum.kShapeEditCmdType,
                  "{" & Me.oInvCLSID.Value & "}",
                  "Prepare part file for Content Center",
                  "Prepare files by defining iProperties and BOMStructure.")

            ' Check if this is the first time the addin has been registered
            If o1stTime Then
                AddToUserInterface()
            End If
        End Sub
#End Region

#Region "UI Customization"
        ' Create Simple Button
        Private Sub CreateSimpleButton(ByVal oDocTypeName As String, ByVal oRibbonTabName As String, ByVal oRibbonPanelName As String, ByVal oButtonDef As ButtonDefinition)
            Dim oRibbonDoc As Ribbon = oInvApp.UserInterfaceManager.Ribbons.Item(oDocTypeName)
            Dim oRibbonTab As RibbonTab = oRibbonDoc.RibbonTabs.Item(oRibbonTabName)
            Dim oRibbonPanel As RibbonPanel = oRibbonTab.RibbonPanels.Item(oRibbonPanelName)
            oRibbonPanel.CommandControls.AddButton(oButtonDef)
        End Sub
#End Region

#Region "User Interface Events"
        ' Sub where the user-interface creation is done. This is called when 
        ' the add-in loaded and also if the user interface is reset. 
        Private Sub AddToUserInterface()
            ' Add buttons to the tools tab of a specified environment. 
            '** Demonstrates creating a button on a new panel of the Tools 
            '** tab of the Part ribbon. 
            ' Get the part ribbon. 
            ' Other available ribbons are Part, Assembly, Drawing, Presentation, and ZeroDoc.

            Dim oAsmPanel As RibbonPanel = Me.oInvApp.UserInterfaceManager.Ribbons.Item("Assembly").RibbonTabs.Item("id_AddInsTab").RibbonPanels.Add("VRepHelper", "id_PanelP_VRep", "{" & Me.oInvCLSID.Value & "}")
            oAsmPanel.CommandControls.AddButton(Me.oVRepRunBtn)

        End Sub
#End Region

#Region "Component Properties"

        Private Sub oVRepRunBtn_OnExecute(ByVal Context As NameValueMap) _
            Handles oVRepRunBtn.OnExecute
            Dim oVRepHelperRun As New VRepHelperRun()
            oVRepHelperRun.Run()

        End Sub

#End Region
    End Class

End Namespace

