Imports System.Data.SqlClient
Imports PuntersEdgeCleanupApp.DatabseActions

Module Module1

    Sub Main()
        Dim db As New DatabseActions

        db.EXECSPROC("Archive_AND_Backup_Data", "")

    End Sub

End Module
