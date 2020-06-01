Imports System
Imports System.Text
Imports System.IO
Imports System.Linq
Imports System.Data
Imports System.Deployment.Application




Public Class Form1

    Public days As Integer
    Public ship As String
    Public voyage As Integer
    Public Twocharshipcode As String
    Public embarkport As String = "1001"
    Public seaday As String = "0"
    Public dateformat As String = "MM/dd/yyyy"
    Public userdesktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Public desktopdi As DirectoryInfo
    Public DebugON As Boolean = False


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call the reset Date Picker and ShipCode & Days

        ResetForm()


    End Sub


    Private Sub ResetForm()
        Populate_Days()
        PopulateShipcodes()
        Reset_Date_Picker()
        Reset_Voyage_Number()
        If ApplicationDeployment.IsNetworkDeployed Then
            lblVersionNumber.Text = "Application Version: " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Else
            lblVersionNumber.Text = "Application Version: " & Application.ProductVersion
        End If
    End Sub

    Private Sub PopulateShipcodes()
        CboShipPicker.Items.Add("Adventure of the Seas")
        CboShipPicker.Items.Add("Allure of the Seas")
        CboShipPicker.Items.Add("Anthem of the Seas")
        CboShipPicker.Items.Add("Brilliance of the Seas")
        CboShipPicker.Items.Add("Enchantment of the Seas")
        CboShipPicker.Items.Add("Explorer of the Seas")
        CboShipPicker.Items.Add("Freedom of the Seas")
        CboShipPicker.Items.Add("Grandeur of the Seas")
        CboShipPicker.Items.Add("Harmony of the Seas")
        CboShipPicker.Items.Add("Independence of the Seas")
        CboShipPicker.Items.Add("Jewel of the Seas")
        CboShipPicker.Items.Add("Liberty of the Seas")
        CboShipPicker.Items.Add("Majesty of the Seas")
        CboShipPicker.Items.Add("Mariner of the Seas")
        CboShipPicker.Items.Add("Navigator of the Seas")
        CboShipPicker.Items.Add("Empress of the Seas")
        CboShipPicker.Items.Add("Oasis of the Seas")
        CboShipPicker.Items.Add("Ovation of the Seas")
        CboShipPicker.Items.Add("Odyssey of the Seas")
        CboShipPicker.Items.Add("Quantum of the Seas")
        CboShipPicker.Items.Add("Radiance of the Seas")
        CboShipPicker.Items.Add("Rhapsody of the Seas")
        CboShipPicker.Items.Add("Serenade of the Seas")
        CboShipPicker.Items.Add("Spectrum of the Seas")
        CboShipPicker.Items.Add("Symphony of the Seas")
        CboShipPicker.Items.Add("Vision of the Seas")
        CboShipPicker.Items.Add("Voyager of the Seas")
        CboShipPicker.Items.Add("Wonder of the Seas")
        CboShipPicker.Items.Add("Ascent")
        CboShipPicker.Items.Add("Apex")
        CboShipPicker.Items.Add("Beyond")
        CboShipPicker.Items.Add("Constellation")
        CboShipPicker.Items.Add("Eclipse")
        CboShipPicker.Items.Add("Edge")
        CboShipPicker.Items.Add("Equinox")
        CboShipPicker.Items.Add("Infinity")
        CboShipPicker.Items.Add("Millennium")
        CboShipPicker.Items.Add("Reflection")
        CboShipPicker.Items.Add("Silhouette")
        CboShipPicker.Items.Add("Solstice")
        CboShipPicker.Items.Add("Summit")

    End Sub

    Private Sub Reset_Voyage_Number()
        txtVoyageNumber.Text = ""
    End Sub

    Private Sub Populate_Days()
        Dim i As Integer
        For i = 1 To 45
            CboNumberOfDays.Items.Add(i)
        Next
        CboNumberOfDays.Text = "0"
    End Sub

    Private Sub Reset_Date_Picker()
        DtpEmbarkation.Value = Now

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Cruisecounter As Integer = 0
        Dim dayofcruise As Integer
        Dim maxdays As Integer
        Dim portcode As Integer
        Dim workingdate As DateTime
        Dim portdate As String
        Dim voyagetovalidate As String
        Dim validatedvoyage As Integer

        'Validate that Cruise is an integer less than 99999
        voyagetovalidate = txtVoyageNumber.Text
        If Integer.TryParse(voyagetovalidate, validatedvoyage) Then
            'Voyage is an integer
            If validatedvoyage > 99999 Then
                MsgBox("Voyage number is not valid")
                Exit Sub
            End If
        Else
            MsgBox("Voyage number is not valid")
            Exit Sub

        End If


        'Set the cruise days from the combo box pushing an error if it is blank.
        If CboNumberOfDays.Text = "" Then
            MsgBox("Set a valid cruise length")
            Exit Sub
        Else
            If Integer.TryParse(CboNumberOfDays.Text, maxdays) Then
                'Correct
            End If

        End If

        If CboShipPicker.Text.ToString = "" Then
            MsgBox("No Ship assigned")
            Exit Sub
        Else
            Twocharshipcode = getshipcode(CboShipPicker.Text.ToString)
        End If


        If DebugON = True Then
            MsgBox(Twocharshipcode)
        End If

        Dim dt As New DataTable

        dt.Columns.Add("ship_code")
        dt.Columns.Add("voyage_num")
        dt.Columns.Add("date")
        dt.Columns.Add("port_id")

        'Set Port to Miami for day 1 and last day

        For dayofcruise = 0 To maxdays
            workingdate = DtpEmbarkation.Value
            workingdate = workingdate.AddDays(dayofcruise)
            If DebugON = True Then
                MsgBox(dayofcruise)
            End If
            portdate = workingdate.ToString(dateformat)
            If DebugON = True Then
                MsgBox(portdate)
            End If
            If dayofcruise = 0 Then
                portcode = embarkport
            ElseIf dayofcruise = maxdays Then
                portcode = embarkport
            Else
                portcode = seaday
            End If

            dt.Rows.Add(Twocharshipcode, validatedvoyage, portdate, portcode)
        Next

        Dim sb As New StringBuilder
        Dim finalcsvfile As String = "Itinerary_" & Twocharshipcode & "_" & validatedvoyage & ".csv"
        Dim fullpath As String = Path.Combine(userdesktop, finalcsvfile)
        desktopdi = New DirectoryInfo(userdesktop)
        If Not desktopdi.Exists Then
            desktopdi.Create()

        End If
        Dim streampath As String = fullpath
        Using writer As StreamWriter = New StreamWriter(streampath)
            Rfc4180Writer.WriteDataTable(dt, writer, True)
        End Using

        MsgBox("Done")
        ResetForm()
    End Sub


    Public Function getshipcode(inputstring)
        If inputstring = "Adventure of the Seas" Then
            Return "AD"
        ElseIf inputstring = "Allure of the Seas" Then
            Return "AL"
        ElseIf inputstring = "Anthem of the Seas" Then
            Return "AN"
        ElseIf inputstring = "Brilliance of the Seas" Then
            Return "BR"
        ElseIf inputstring = "Enchantment of the Seas" Then
            Return "EN"
        ElseIf inputstring = "Explorer of the Seas" Then
            Return "EX"
        ElseIf inputstring = "Freedom of the Seas" Then
            Return "FR"
        ElseIf inputstring = "Grandeur of the Seas" Then
            Return "GR"
        ElseIf inputstring = "Harmony of the Seas" Then
            Return "HM"
        ElseIf inputstring = "Independence of the Seas" Then
            Return "ID"
        ElseIf inputstring = "Jewel of the Seas" Then
            Return "JW"
        ElseIf inputstring = "Liberty of the Seas" Then
            Return "LB"
        ElseIf inputstring = "Majesty of the Seas" Then
            Return "MJ"
        ElseIf inputstring = "Mariner of the Seas" Then
            Return "MA"
        ElseIf inputstring = "Navigator of the Seas" Then
            Return "NV"
        ElseIf inputstring = "Empress of the Seas" Then
            Return "NE"
        ElseIf inputstring = "Oasis of the Seas" Then
            Return "OA"
        ElseIf inputstring = "Ovation of the Seas" Then
            Return "OV"
        ElseIf inputstring = "Odyssey of the Seas" Then
            Return "OY"
        ElseIf inputstring = "Quantum of the Seas" Then
            Return "QN"
        ElseIf inputstring = "Radiance of the Seas" Then
            Return "RD"
        ElseIf inputstring = "Rhapsody of the Seas" Then
            Return "RH"
        ElseIf inputstring = "Serenade of the Seas" Then
            Return "SR"
        ElseIf inputstring = "Spectrum of the Seas" Then
            Return "SC"
        ElseIf inputstring = "Symphony of the Seas" Then
            Return "SY"
        ElseIf inputstring = "Vision of the Seas" Then
            Return "VI"
        ElseIf inputstring = "Voyager of the Seas" Then
            Return "VY"
        ElseIf inputstring = "Wonder of the Seas" Then
            Return "WN"
        ElseIf inputstring = "Ascent" Then
            Return "AS"
        ElseIf inputstring = "Apex" Then
            Return "AX"
        ElseIf inputstring = "Beyond" Then
            Return "BY"
        ElseIf inputstring = "Constellation" Then
            Return "CS"
        ElseIf inputstring = "Eclipse" Then
            Return "EC"
        ElseIf inputstring = "Edge" Then
            Return "EG"
        ElseIf inputstring = "Equinox" Then
            Return "EQ"
        ElseIf inputstring = "Infinity" Then
            Return "IN"
        ElseIf inputstring = "Millennium" Then
            Return "ML"
        ElseIf inputstring = "Reflection" Then
            Return "RF"
        ElseIf inputstring = "Silhouette" Then
            Return "SI"
        ElseIf inputstring = "Solstice" Then
            Return "SL"
        ElseIf inputstring = "Summit" Then
            Return "SM"
        Else
            Return ""
        End If

    End Function



    Public Class Rfc4180Writer
        Public Shared Sub WriteDataTable(ByVal sourceTable As DataTable,
            ByVal writer As TextWriter, ByVal includeHeaders As Boolean)
            If (includeHeaders) Then
                Dim headerValues As IEnumerable(Of String) = sourceTable.Columns.OfType(Of DataColumn).Select(Function(column) QuoteValue(column.ColumnName))

                writer.WriteLine(String.Join(",", headerValues))
            End If

            Dim items As IEnumerable(Of String) = Nothing
            For Each row As DataRow In sourceTable.Rows
                items = row.ItemArray.Select(Function(obj) QuoteValue(If(obj?.ToString(), String.Empty)))
                writer.WriteLine(String.Join(",", items))
            Next

            writer.Flush()
        End Sub

        Private Shared Function QuoteValue(ByVal value As String) As String
            Return String.Concat("""", value.Replace("""", """"""), """")
        End Function
    End Class

End Class


