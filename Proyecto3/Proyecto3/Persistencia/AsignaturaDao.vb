Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class AsignaturaDao
    ' Atributos de la clase gestorAsignatura ' 
    Private mlistaAsignatura As Collection
    Private Agente As AgenteBD
    'Constructor ' 

    Sub New()
        mlistaAsignatura = New Collection
    End Sub

    Public Property listaAsignatura As Collection
        Get
            Return listaAsignatura
        End Get
        Set(value As Collection)

        End Set
    End Property

    ' Iserta una persona y Asignaturas en la base de datos ' 
    Sub insert(ByVal newAsignatura As Asignatura)
        Agente = AgenteBD.getInstancia()
        Dim aSQL As String = "INSERT INTO ASIGNATURA VALUES ('" + newAsignatura.ID + "');" 'Querie'
        Agente.create(aSQL)
    End Sub

    ' Actualizar una persona y Asignaturas en la base de datos ' 
    Sub update(ByVal newAsignatura As Asignatura)
        Agente = AgenteBD.getInstancia()
        Dim aSQL As String = "UPDATE ASIGNATURA SET '" + newAsignatura.ID + "';"
        Agente.update(aSQL)
    End Sub


    'Borrar la persona y Asignatura elegida de la base de datos'
    Sub delete(ByVal noAsignatura As Asignatura)
        Agente = AgenteBD.getInstancia()
        Dim aSQL As String = "DELETE * FROM ASIGNATURA WHERE ID='" + noAsignatura.ID + "';"
        Agente.delete(aSQL)
    End Sub

    'Borrar todas las personas y Asignaturas de la base de datos'
    Sub deleteAll(ByVal noAsignatura As Asignatura)
        Agente = AgenteBD.getInstancia()
        Dim aSQL As String = "DELETE * FROM ASIGNATURA;"
        Agente.delete(aSQL)
    End Sub

    Friend Sub readAll()
        Agente = AgenteBD.getInstancia
        Dim aSQL As String = "SELECT NOMBRE FROM ASIGNATURA;"
        Dim command As New SqlCommand()
        Dim reader As SqlDataReader = command.ExecuteReader()
        While reader.Read()
            Dim elemento As String = reader("NOMBRE").ToString()
            Form1.Lista_Personas.Items.Add(elemento)
            reader.Close()
        End While
    End Sub
End Class
