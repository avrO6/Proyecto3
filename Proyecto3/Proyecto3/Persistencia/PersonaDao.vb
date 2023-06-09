Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class PersonaDao
    ' Atributos de la clase gestorPersona ' 
    Private mlistaPersona As Collection
    Private Agente As AgenteBD
    'Constructor ' 
    Sub New()
        mlistaPersona = New Collection
    End Sub

    Public Property listaPersona As Collection
        Get
            Return listaPersona
        End Get
        Set(value As Collection)

        End Set
    End Property

    ' Iserta una persona y Asignaturas en la base de datos ' 
    Sub insert(ByVal newPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim pSQL As String = "INSERT INTO PERSONA VALUES ('" + newPersona.NDNI1 + "', '" + newPersona.NOM + "');"   'Cabiar line de insert'
        Agente.create(pSQL)

    End Sub

    ' Actualizar una persona en la base de datos ' 
    Sub update(ByVal newPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim pSQL As String = "UPDATE PERSONA SET '" + newPersona.NDNI1 + " AND " + newPersona.NOM + "';"  'cambiar update'
        Agente.update(pSQL)
    End Sub

    'Borrar la persona elegida de la base de datos'
    Sub delete(ByVal noPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim pSQL As String = "DELETE FROM PERSONA WHERE DNI='" + noPersona.NDNI1 + "';"
        Agente.delete(pSQL)
    End Sub

    'Borrar todas las personas de la base de datos'
    Sub deleteAll(ByVal noPersona As Persona)
        Agente = AgenteBD.getInstancia()
        Dim pSQL As String = "DELETE * FROM PERSONA;"
        Agente.delete(pSQL)
    End Sub

    Friend Sub readAll()
        Agente = AgenteBD.getInstancia
        Dim aSQL As String = "SELECT NOMBRE FROM PERSONA;"
        Dim command As New SqlCommand()
        Dim reader As SqlDataReader = command.ExecuteReader()
        While reader.Read()
            Dim elemento As String = reader("NOMBRE").ToString()
            Form1.Lista_Personas.Items.Add(elemento)
            reader.Close()
        End While
    End Sub
End Class
