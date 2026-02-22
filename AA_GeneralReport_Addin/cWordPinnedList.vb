Imports System.Xml.Linq
Public Class cWordPinnedList
    Public Sub New()

    End Sub
    '
    Public Sub pin_Remove_TemplateFromNewList(templatePath As String)

        Dim xmlPath As String = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Microsoft\Templates\TemplateProperties.xml"
    )

        If Not System.IO.File.Exists(xmlPath) Then
            Exit Sub
        End If

        Dim doc As XDocument = XDocument.Load(xmlPath)
        Dim ns = doc.Root.GetDefaultNamespace()

        ' Find all <template> nodes with matching path
        Dim toRemove = doc.Descendants(ns + "template").
        Where(Function(t)
                  Dim p = t.Attribute("path")
                  Return p IsNot Nothing AndAlso
                         String.Equals(p.Value, templatePath, StringComparison.OrdinalIgnoreCase)
              End Function).
        ToList()

        ' Remove them
        For Each t In toRemove
            t.Remove()
        Next

        ' Save the updated XML
        doc.Save(xmlPath)

    End Sub


End Class
