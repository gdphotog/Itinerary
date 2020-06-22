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
    Public StaffName As String = "COVID 19, Test User"
    Public Staffid As Integer = 10000
    Public cruiseline As String
    Public DebugON As Boolean = False
    Public firstrun As Boolean = True


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call the reset Date Picker and ShipCode & Days


        If firstrun = True Then
            Setupform()
            firstrun = False

        Else
        ResetForm()
        End If

    End Sub
    Private Sub setupform()
        Populate_Days()
        PopulateShipcodes()
        If ApplicationDeployment.IsNetworkDeployed Then
            lblVersionNumber.Text = "Application Version: " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Else
            lblVersionNumber.Text = "Application Version: " & Application.ProductVersion
        End If

    End Sub

    Private Sub ResetForm()
        Reset_ship()
        Reset_days()
        Reset_Date_Picker()
        Reset_Voyage_Number()

    End Sub

    Private Sub PopulateShipcodes()
        CboShipPicker.Items.Add(" ")
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
        For i = 0 To 45
            CboNumberOfDays.Items.Add(i)
        Next
        CboNumberOfDays.Text = "0"
    End Sub

    Private Sub Reset_ship()
        CboNumberOfDays.Text = " "
    End Sub
    Private Sub Reset_days()
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
        If CboNumberOfDays.Text = "0" Then
            MsgBox("Set a valid cruise length")
            Exit Sub
        Else
            If Integer.TryParse(CboNumberOfDays.Text, maxdays) Then
                'Correct
            End If

        End If

        If CboShipPicker.Text.ToString = " " Then
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
        'Iterates through the length of the cruise amd sets the ports

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

            'Add the row to the Datatable

            dt.Rows.Add(Twocharshipcode, validatedvoyage, portdate, portcode)
        Next

        'Write out the Array to a stringbuilder format
        Dim sb As New StringBuilder
        Dim finalItineraryfile As String = "Itinerary_" & Twocharshipcode & "_" & validatedvoyage & ".csv"
        Dim fullpath As String = Path.Combine(userdesktop, finalItineraryfile)

        'Checks the Desktop folder exists
        desktopdi = New DirectoryInfo(userdesktop)
        If Not desktopdi.Exists Then
            desktopdi.Create()

        End If
        Dim itinerarystreampath As String = fullpath

        'Writes the Stringbuilder to file as a stream. The CSV file is RFC4180 compliant.
        'https://tools.ietf.org/html/rfc4180

        Using writer As StreamWriter = New StreamWriter(itinerarystreampath)
            Rfc4180Writer.WriteDataTable(dt, writer, True)
        End Using

        'Sets the variables for the Staff and Photogtype for the Staff file.
        Dim photogtype As String
        Dim staffdt As New DataTable

        'Creates the datatable format for the Staff CSV file
        staffdt.Columns.Add("ship_code")
        staffdt.Columns.Add("voyage_num")
        staffdt.Columns.Add("photog_id")
        staffdt.Columns.Add("photog_type")
        staffdt.Columns.Add("name_prefix")
        staffdt.Columns.Add("name")

        If cruiseline = "RCI" Then
            'Sets the Photog Type to Lab Tech for Royal
            photogtype = "RPLT"
        ElseIf cruiseline = "CEL" Then
            'Sets the Photog Type to Lab Tech for CEL.
            photogtype = "CPLT"
        End If

        staffdt.Rows.Add(Twocharshipcode, validatedvoyage, Staffid, photogtype, "Mr", StaffName)

        Dim staffsb As New StringBuilder
        Dim finalStafffile As String = "Staff_" & Twocharshipcode & "_" & validatedvoyage & ".csv"
        Dim stafffullpath As String = Path.Combine(userdesktop, finalStafffile)

        Dim staffstreampath As String = stafffullpath

        'Writes the Stringbuilder to file as a stream. The CSV file is RFC4180 compliant.
        'https://tools.ietf.org/html/rfc4180

        Using writer As StreamWriter = New StreamWriter(staffstreampath)
            Rfc4180Writer.WriteDataTable(staffdt, writer, True)
        End Using

        MsgBox("Done")
        ResetForm()
    End Sub


    Public Function getshipcode(inputstring)
        If inputstring = "Adventure of the Seas" Then
            cruiseline = "RCI"
            Return "AD"
        ElseIf inputstring = "Allure of the Seas" Then
            cruiseline = "RCI"
            Return "AL"
        ElseIf inputstring = "Anthem of the Seas" Then
            cruiseline = "RCI"
            Return "AN"
        ElseIf inputstring = "Brilliance of the Seas" Then
            cruiseline = "RCI"
            Return "BR"
        ElseIf inputstring = "Enchantment of the Seas" Then
            cruiseline = "RCI"
            Return "EN"
        ElseIf inputstring = "Explorer of the Seas" Then
            cruiseline = "RCI"
            Return "EX"
        ElseIf inputstring = "Freedom of the Seas" Then
            cruiseline = "RCI"
            Return "FR"
        ElseIf inputstring = "Grandeur of the Seas" Then
            cruiseline = "RCI"
            Return "GR"
        ElseIf inputstring = "Harmony of the Seas" Then
            cruiseline = "RCI"
            Return "HM"
        ElseIf inputstring = "Independence of the Seas" Then
            cruiseline = "RCI"
            Return "ID"
        ElseIf inputstring = "Jewel of the Seas" Then
            cruiseline = "RCI"
            Return "JW"
        ElseIf inputstring = "Liberty of the Seas" Then
            cruiseline = "RCI"
            Return "LB"
        ElseIf inputstring = "Majesty of the Seas" Then
            cruiseline = "RCI"
            Return "MJ"
        ElseIf inputstring = "Mariner of the Seas" Then
            cruiseline = "RCI"
            Return "MA"
        ElseIf inputstring = "Navigator of the Seas" Then
            cruiseline = "RCI"
            Return "NV"
        ElseIf inputstring = "Empress of the Seas" Then
            cruiseline = "RCI"
            Return "NE"
        ElseIf inputstring = "Oasis of the Seas" Then
            cruiseline = "RCI"
            Return "OA"
        ElseIf inputstring = "Ovation of the Seas" Then
            cruiseline = "RCI"
            Return "OV"
        ElseIf inputstring = "Odyssey of the Seas" Then
            cruiseline = "RCI"
            Return "OY"
        ElseIf inputstring = "Quantum of the Seas" Then
            cruiseline = "RCI"
            Return "QN"
        ElseIf inputstring = "Radiance of the Seas" Then
            cruiseline = "RCI"
            Return "RD"
        ElseIf inputstring = "Rhapsody of the Seas" Then
            cruiseline = "RCI"
            Return "RH"
        ElseIf inputstring = "Serenade of the Seas" Then
            cruiseline = "RCI"
            Return "SR"
        ElseIf inputstring = "Spectrum of the Seas" Then
            cruiseline = "RCI"
            Return "SC"
        ElseIf inputstring = "Symphony of the Seas" Then
            cruiseline = "RCI"
            Return "SY"
        ElseIf inputstring = "Vision of the Seas" Then
            cruiseline = "RCI"
            Return "VI"
        ElseIf inputstring = "Voyager of the Seas" Then
            cruiseline = "RCI"
            Return "VY"
        ElseIf inputstring = "Wonder of the Seas" Then
            cruiseline = "RCI"
            Return "WN"
        ElseIf inputstring = "Ascent" Then
            cruiseline = "CEL"
            Return "AS"
        ElseIf inputstring = "Apex" Then
            cruiseline = "CEL"
            Return "AX"
        ElseIf inputstring = "Beyond" Then
            cruiseline = "CEL"
            Return "BY"
        ElseIf inputstring = "Constellation" Then
            cruiseline = "CEL"
            Return "CS"
        ElseIf inputstring = "Eclipse" Then
            cruiseline = "CEL"
            Return "EC"
        ElseIf inputstring = "Edge" Then
            cruiseline = "CEL"
            Return "EG"
        ElseIf inputstring = "Equinox" Then
            cruiseline = "CEL"
            Return "EQ"
        ElseIf inputstring = "Infinity" Then
            cruiseline = "CEL"
            Return "IN"
        ElseIf inputstring = "Millennium" Then
            cruiseline = "CEL"
            Return "ML"
        ElseIf inputstring = "Reflection" Then
            cruiseline = "CEL"
            Return "RF"
        ElseIf inputstring = "Silhouette" Then
            cruiseline = "CEL"
            Return "SI"
        ElseIf inputstring = "Solstice" Then
            cruiseline = "CEL"
            Return "SL"
        ElseIf inputstring = "Summit" Then
            cruiseline = "CEL"
            Return "SM"
        Else
            cruiseline = ""
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


