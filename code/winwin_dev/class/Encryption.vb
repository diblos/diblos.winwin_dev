Imports System.Windows.Forms

Public Class Encryption
    Public Sub ProtectConnectionString()
        ToggleConnectionStringProtection(Application.ExecutablePath, True)
    End Sub

    Public Sub ProtectApplicationSetting()
        ToggleApplicationSettingProtection(Application.ExecutablePath, True)
    End Sub

    Public Sub UnprotectConnectionString()
        ToggleConnectionStringProtection(Application.ExecutablePath, False)
    End Sub

    Public Sub UnprotectApplicationSetting()
        ToggleApplicationSettingProtection(Application.ExecutablePath, False)
    End Sub

    Private Shared Sub ToggleConnectionStringProtection(ByVal pathName As String, ByVal protect As Boolean)
        ' Define the Dpapi provider name.
        Dim strProvider As String = "DataProtectionConfigurationProvider"
        'Dim strProvider As String = "RSAProtectedConfigurationProvider"

        Dim oConfiguration As System.Configuration.Configuration = Nothing
        Dim oSection As System.Configuration.ConnectionStringsSection = Nothing

        Try
            ' Open the configuration file and retrieve the connectionStrings section.

            ' For Web!
            ' oConfiguration = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");

            ' For Windows!
            ' Takes the executable file name without the config extension.
            oConfiguration = System.Configuration.ConfigurationManager.OpenExeConfiguration(pathName)

            If oConfiguration IsNot Nothing Then
                Dim blnChanged As Boolean = False
                oSection = TryCast(oConfiguration.GetSection("connectionStrings"), System.Configuration.ConnectionStringsSection)

                If oSection IsNot Nothing Then
                    If (Not (oSection.ElementInformation.IsLocked)) AndAlso (Not (oSection.SectionInformation.IsLocked)) Then
                        If protect Then
                            If Not (oSection.SectionInformation.IsProtected) Then
                                blnChanged = True

                                ' Encrypt the section.
                                oSection.SectionInformation.ProtectSection(strProvider)
                            End If
                        Else
                            If oSection.SectionInformation.IsProtected Then
                                blnChanged = True

                                ' Remove encryption.
                                oSection.SectionInformation.UnprotectSection()
                            End If
                        End If
                    End If

                    If blnChanged Then
                        ' Indicates whether the associated configuration section will be saved even if it has not been modified.
                        oSection.SectionInformation.ForceSave = True

                        ' Save the current configuration.
                        oConfiguration.Save()
                    End If
                End If
            End If
        Catch ex As System.Exception
            Throw (ex)
        Finally
        End Try
    End Sub

    Private Shared Sub ToggleApplicationSettingProtection(ByVal pathName As String, ByVal protect As Boolean)
        ' Define the Dpapi provider name.
        Dim strProvider As String = "DataProtectionConfigurationProvider"
        'Dim strProvider As String = "RSAProtectedConfigurationProvider"

        Dim oConfiguration As System.Configuration.Configuration = Nothing
        Dim oSection As System.Configuration.AppSettingsSection = Nothing

        Try
            ' Open the configuration file and retrieve the connectionStrings section.

            ' For Web!
            ' oConfiguration = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");

            ' For Windows!
            ' Takes the executable file name without the config extension.
            oConfiguration = System.Configuration.ConfigurationManager.OpenExeConfiguration(pathName)

            If oConfiguration IsNot Nothing Then
                Dim blnChanged As Boolean = False
                oSection = TryCast(oConfiguration.GetSection("appSettings"), System.Configuration.AppSettingsSection)

                If oSection IsNot Nothing Then
                    If (Not (oSection.ElementInformation.IsLocked)) AndAlso (Not (oSection.SectionInformation.IsLocked)) Then
                        If protect Then
                            If Not (oSection.SectionInformation.IsProtected) Then
                                blnChanged = True

                                ' Encrypt the section.
                                oSection.SectionInformation.ProtectSection(strProvider)
                            End If
                        Else
                            If oSection.SectionInformation.IsProtected Then
                                blnChanged = True

                                ' Remove encryption.
                                oSection.SectionInformation.UnprotectSection()
                            End If
                        End If
                    End If

                    If blnChanged Then
                        ' Indicates whether the associated configuration section will be saved even if it has not been modified.
                        oSection.SectionInformation.ForceSave = True

                        ' Save the current configuration.
                        oConfiguration.Save()
                    End If
                End If
            End If
        Catch ex As System.Exception
            Throw (ex)
        Finally
        End Try
    End Sub
End Class
