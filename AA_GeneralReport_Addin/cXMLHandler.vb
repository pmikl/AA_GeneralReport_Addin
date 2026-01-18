Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.CustomProperties
Imports DocumentFormat.OpenXml
Imports System.IO
Public Class cXMLHandler
    '
    Public Sub New()

    End Sub
    '

    Sub RemoveAssemblyLocationProperty(docPath As String)
        If Not System.IO.File.Exists(docPath) Then
            Throw New FileNotFoundException("Document not found.", docPath)
        End If

        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(docPath, True)
            Dim customPropsPart = wordDoc.CustomFilePropertiesPart
            If customPropsPart IsNot Nothing Then
                Dim props = customPropsPart.Properties
                If props IsNot Nothing Then
                    Dim targetProp = props.Elements(Of CustomDocumentProperty)().
                    FirstOrDefault(Function(p) p.Name.HasValue AndAlso p.Name.Value = "_AssemblyLocation")

                    If targetProp IsNot Nothing Then
                        targetProp.Remove()
                        props.Save()
                        Console.WriteLine("Property '_AssemblyLocation' removed.")
                    Else
                        Console.WriteLine("Property '_AssemblyLocation' not found.")
                    End If
                End If
            End If
        End Using
    End Sub


End Class
