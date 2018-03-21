Imports System.Data.OracleClient
Public Class ModVariables
    'Public Con As New SqlConnection("User Id=sa;Database=NorthWind;Server=(local)")
    Public Shared Usuario, Apellido, Nombre As String
    Public Enum Operaciones
        Adicionar = 0
        Actualizar = 1
        Eliminar = 2
    End Enum
    Public Operacion As Operaciones
End Class
