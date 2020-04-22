Imports System.Runtime.CompilerServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module Json

    <Extension()>
    Public Function From_JSON(ByVal s As String) As Dictionary(Of String, Object)
        Return JsonConvert.DeserializeObject(s, GetType(Dictionary(Of String, Object)))
    End Function

    <Extension()>
    Public Function From_JSON(ByVal s As String, ByVal ObjType As Type) As Object
        Return JsonConvert.DeserializeObject(s, ObjType)
    End Function

    <Extension()>
    Public Function To_JSON(ByVal o As Object, Optional ByVal f As Formatting = Formatting.Indented) As String
        Dim nj As String = JsonConvert.SerializeObject(o, f)

        If TryCast(o, IEnumerable) Is Nothing Then
            ' object - sort
            Dim parsedObject = JObject.Parse(nj)
            Dim normalizedObject = JSON_SortPropertiesAlphabetically(parsedObject)
            Return JsonConvert.SerializeObject(normalizedObject, f)
        Else
            ' list - don't sort
            Return nj
        End If
    End Function

    <Extension()>
    Public Function To_JSON_SelectToken(ByVal o As Object, ByVal Token As String) As JToken
        Dim obj = JObject.FromObject(o)
        Return obj.SelectToken(Token)
    End Function

    Private Function JSON_SortPropertiesAlphabetically(ByVal original As JObject) As JObject
        Dim result = New JObject()

        For Each [property] In original.Properties().ToList().OrderBy(Function(p) p.Name)
            Dim value = TryCast([property].Value, JObject)

            If value IsNot Nothing Then
                value = JSON_SortPropertiesAlphabetically(value)
                result.Add([property].Name, value)
            Else
                result.Add([property].Name, [property].Value)
            End If
        Next

        Return result
    End Function

    Private Function JSON_DeserialiseFromStream(Of t)(ByVal s As IO.Stream) As t
        ' adapted from: https://stackify.com/improved-performance-using-json-streaming/

        If s Is Nothing Then
            ' err
        End If

        If s.CanRead = False Then
            ' err
        End If

        Dim DeserialisedObj As t

        Dim f = GetType(t)

        Using sr As New IO.StreamReader(s)
            Using reader As Newtonsoft.Json.JsonReader = New Newtonsoft.Json.JsonTextReader(sr)
                Dim Serialiiser As New Newtonsoft.Json.JsonSerializer()
                DeserialisedObj = Serialiiser.Deserialize(Of t)(reader)
            End Using
        End Using

        Return DeserialisedObj

    End Function

    Public Function JSON_DeserialiseFromFile(Of t)(ByVal Path As String) As t

        Dim DeserialisedObj As t

        Dim f = GetType(t)

        Using fs As IO.FileStream = New IO.FileStream(Path, IO.FileMode.Open, IO.FileAccess.Read)
            DeserialisedObj = JSON_DeserialiseFromStream(Of t)(fs)
        End Using

        Return DeserialisedObj

    End Function

End Module
