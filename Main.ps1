#test
Sub Tome()
Dossier = InputBox("dossier") & "\"
    tbNom = Split(Dossier, "\")
    strNom = tbNom(UBound(tbNom) - 1)

    Dim StrFile As String
    StrFile = Dir(Dossier)
    Do While Len(StrFile) > 0
        oldName = Dossier & "\" & StrFile
        Extension = Right(oldName, 4)
        
        Dim i As Byte, Nb As Byte
        Dim Cible As String, Resultat As String
        Dim Nombre As Double
        
        Cible = StrFile
        'Pour que fonction Val puisse reconnaitre les décimales: Remplacement des
        'virgules par des points
        Cible = Replace(Cible, ",", ".")
        'Pour gérer deux nombres qui se suivent: remplacement des espaces
        'par un caractère Alpha
        Cible = Replace(Cible, " ", "x")
        cc = 0
        
        For i = 1 To Len(Cible)
            If IsNumeric(Mid(Cible, i, 1)) Then
                If cc = 0 Then
                Nombre = Val(Mid(Cible, i, Len(Cible) - i + 1))
                Exit For
                Else
                cc = cc + 1
                End If
            End If
        
        Next
                
        Select Case Nombre
            Case Is < 10
            strNombre = "00" & Nombre
            Case Is < 100
            strNombre = "0" & Nombre
            Case Else
            strNombre = Nombre
        
        End Select
        
'        newName = Replace(oldName, strDel, "")
        newName = Dossier & "\" & strNom & " Vol." & strNombre & Extension  'Vol.TBD Ch. '
        Name oldName As newName ' Rename file.
        StrFile = Dir
    Loop
End Sub