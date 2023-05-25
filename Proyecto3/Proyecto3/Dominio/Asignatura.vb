Imports System.Data.OleDb

Public Class Asignatura

    'Atributos de la clase Asignatura'
    Private mId As String
    Private mNombre As String
    Private mCurso As String
    Private mHoras As Date
    Private Gestor As AsignaturaDao

    ' Constructores '
    Public Sub New()

    End Sub

    Public Sub New(mId As String, mNombre As String, mCurso As String, mHoras As String)
        Me.mId = mId
        Me.mNombre = mNombre
        Me.mHoras = mHoras
        Me.mCurso = mCurso

    End Sub

    'Propertys'
    Public Property ID As String
        Get
            Return mId
        End Get
        Set(value As String)
            mId = value
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

    Public Property Curso As String
        Get
            Return mCurso
        End Get
        Set(value As String)
            mCurso = value
        End Set
    End Property

    Public Property Horas As String
        Get
            Return mHoras
        End Get
        Set(value As String)
            mHoras = value
        End Set
    End Property


    ' Funcion para leer todos ' 
    Function leertodos()
        Dim todos As Collection
        Me.Gestor.readAll()
        todos = Gestor.listaAsignatura
        Return todos
    End Function

    ' Metodo insertar Asignatura '
    Sub insertarPersona()
        Me.Gestor.insert(Me)
    End Sub

    ' Metodo modificar Asignatura ' 
    Sub modificarPersona()
        Me.Gestor.update(Me)
    End Sub

End Class
