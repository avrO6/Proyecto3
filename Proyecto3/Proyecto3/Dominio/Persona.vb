Imports System.Data.OleDb


Public Class Persona

    'Atributos de la clase Persona'
    Private mDNI As String
    Private mNombre As String
    Private Gestor As PersonaDao

    ' Constructor '
    Sub New()
        Me.mDNI = mDNI
        Me.mNombre = mNombre
        Gestor = New PersonaDao
    End Sub

    'Propertys'
    Public Property NDNI1 As String
        Get
            Return mDNI
        End Get
        Set(value As String)
            mDNI = value
        End Set
    End Property

    Public Property NOM As String
        Get
            Return mNombre
        End Get
        Set(value As String)
            mNombre = value
        End Set
    End Property

    ' Funcion para leer todos ' 
    Function leertodos()
        Dim todos As Collection
        Me.Gestor.readAll()
        todos = Gestor.listaPersona
        Return todos
    End Function

    ' Metodo insertar Persona '
    Sub insertarPersona()
        Me.Gestor.insert(Me)
    End Sub

    ' Metodo modificar Persona ' 
    Sub modificarPersona()
        Me.Gestor.update(Me)
    End Sub

    Function leerpersona(text As String)
        If NDNI1 IsNot Nothing Then
            MsgBox("La persona ya existe")
        End If
    End Function

    Sub eliminarPersona()
        Me.Gestor.delete(Me)
    End Sub

End Class