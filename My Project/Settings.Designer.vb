' ***********************************************************************
' Assembly         : AutoStoreLibrary
' Author           : John Campbell-Higgens
' Created          : 16-Sep-2020
'
' Last Modified By : John Campbell-Higgens
' Last Modified On : 06-Jan-2021
' ***********************************************************************
' <copyright file="Settings.Designer.vb" company="John Campbell-Higgens (Kofax UK Ltd)">
'     Copyright © 2020 John Campbell-Higgens (Kofax UK Ltd)
' </copyright>
' <summary></summary>
' ***********************************************************************


Option Strict On
Option Explicit On




Namespace My

    ''' <summary>
    ''' Class MySettings. This class cannot be inherited.
    ''' Implements the <see cref="System.Configuration.ApplicationSettingsBase" />
    ''' </summary>
    ''' <seealso cref="System.Configuration.ApplicationSettingsBase" />
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.7.0.0"),
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase

        ''' <summary>
        ''' The default instance
        ''' </summary>
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()), MySettings)

#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region

        ''' <summary>
        ''' Gets the default.
        ''' </summary>
        ''' <value>The default.</value>
        Public Shared ReadOnly Property [Default]() As MySettings
            Get

#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
    End Class
End Namespace



Namespace My

    ''' <summary>
    ''' Class MySettingsProperty.
    ''' </summary>
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>
    Friend Module MySettingsProperty

        ''' <summary>
        ''' Gets the settings.
        ''' </summary>
        ''' <value>The settings.</value>
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>
        Friend ReadOnly Property Settings() As Global.AutoStoreLibrary.My.MySettings
            Get
                Return Global.AutoStoreLibrary.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
