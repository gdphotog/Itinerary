
Imports System.Text
Imports System.IO
Imports System.Deployment.Application




Public Class Form1

    'AppWide Variable declarations

    'Variables for the folders 
    Public userdesktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Public desktopdi As DirectoryInfo

    'Variables for the setup of the application
    Public DebugON As Boolean
    Public firstrun As Boolean = True
    Public numberofguests As Integer
    Public runmode As String
    Public manifest As Boolean


    'Variables for the app
    Public days As Integer
    Public ship As String
    Public voyage As Integer
    Public Twocharshipcode As String
    Public embarkport As String = "1001"
    Public seaday As String = "0"
    Public dateformat As String = "MM/dd/yyyy"
    Public StaffName As String = "COVID 19, Test User"
    Public Staffid As Integer = 10000
    Public cruiseline As String


    'Variables for the conversion of number to words
    Public mOnesArray(8) As String
    Public mOneTensArray(9) As String
    Public mTensArray(7) As String
    Public mPlaceValues(4) As String

    'variable for Random numbers. It is declared here as if it was done in the function it returns the same numbers each time due to the way that
    'VB.NET uses the time to generate random numbers.
    Dim random As New Random()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'Call the reset Date Picker and ShipCode & Days


    If firstrun = True Then

      'Read settings from app config file
      Try
        DebugON = Convert.ToBoolean(My.Settings.DebugOn)
        runmode = My.Settings.RunMode.ToUpper

        'Use CASE for the runmode setting
        Select Case runmode

          Case "SHIP"
            manifest = False

          Case "LAB"
            manifest = True

          Case "HOME"
            manifest = True

          Case Else
            Throw New Exception("RunMode not set correctly in application Config file")

        End Select
      Catch ex As Exception
        MsgBox(ex.Message, MsgBoxStyle.Critical)
        End
      End Try

      'If Manifest is needed then parse to see if Guest is integer in AppConfig file.
      If manifest = True Then
        Try
          numberofguests = Integer.Parse(My.Settings.NumberOfGuests)
        Catch ex As Exception
          MsgBox(ex.Message, MsgBoxStyle.Critical, "Number of Guests set incorrectly in App Config")
          End
        End Try
      End If

      'If Debug is enabled show message box
      If DebugON = True Then
        MsgBox("The current run mode is : " & runmode)
      End If
      Setupform()
      firstrun = False

    Else
      ResetForm()
        End If

    End Sub
    Private Sub Setupform()
        populatewords()
        PopulateShipcodes()
        Populate_Days()

    If ApplicationDeployment.IsNetworkDeployed Then
      If runmode <> "SHIP" Then
        lblVersionNumber.Text = "Application Version: " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString & "     RunMode: " & runmode
      Else
        lblVersionNumber.Text = "Application Version: " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
      End If
    Else
      If runmode <> "SHIP" Then
        lblVersionNumber.Text = "Application Version: " & Application.ProductVersion & "     RunMode: " & runmode
      Else
        lblVersionNumber.Text = "Application Version: " & Application.ProductVersion
      End If

    End If

    End Sub

    Private Sub populatewords()

        mOnesArray(0) = "one"
        mOnesArray(1) = "two"
        mOnesArray(2) = "three"
        mOnesArray(3) = "four"
        mOnesArray(4) = "five"
        mOnesArray(5) = "six"
        mOnesArray(6) = "seven"
        mOnesArray(7) = "eight"
        mOnesArray(8) = "nine"

        mOneTensArray(0) = "ten"
        mOneTensArray(1) = "eleven"
        mOneTensArray(2) = "twelve"
        mOneTensArray(3) = "thirteen"
        mOneTensArray(4) = "fourteen"
        mOneTensArray(5) = "fifteen"
        mOneTensArray(6) = "sixteen"
        mOneTensArray(7) = "seventeen"
        mOneTensArray(8) = "eightteen"
        mOneTensArray(9) = "nineteen"

        mTensArray(0) = "twenty"
        mTensArray(1) = "thirty"
        mTensArray(2) = "forty"
        mTensArray(3) = "fifty"
        mTensArray(4) = "sixty"
        mTensArray(5) = "seventy"
        mTensArray(6) = "eighty"
        mTensArray(7) = "ninety"

        mPlaceValues(0) = "hundred"
        mPlaceValues(1) = "thousand"
        mPlaceValues(2) = "million"
        mPlaceValues(3) = "billion"
        mPlaceValues(4) = "trillion"

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
        Dim populatedays As Integer
        For populatedays = 0 To 45

            CboNumberOfDays.Items.Add(populatedays)
        Next
        CboNumberOfDays.Text = "0"
    End Sub

    Private Sub Reset_ship()
        CboNumberOfDays.Text = " "
    End Sub
    Private Sub Reset_days()
        CboNumberOfDays.Text = "0"
        lblCalcEndDate.Text = " "
    End Sub
    Private Sub Reset_Date_Picker()
        DtpEmbarkation.Value = Now
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnCreateFile.Click

        Dim Cruisecounter As Integer = 0
        Dim dayofcruise As Integer
        Dim maxdays As Integer
        Dim portcode As Integer
        Dim workingdate As DateTime
        Dim portdate As String
        Dim voyagetovalidate As String
        Dim validatedvoyage As Integer
        Dim manifestStartdate As String = ""
        Dim manifestenddate As String = ""


        'Validate that Cruise is an integer less than 99999
        voyagetovalidate = txtVoyageNumber.Text
        If Integer.TryParse(voyagetovalidate, validatedvoyage) Then
            'Voyage is an integer
            If validatedvoyage > 99999 Then
                MsgBox("Voyage number is not valid")
                Exit Sub
            End If
        Else
            MsgBox("Voyage number is not valid. It must be numerical only")
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
                manifestStartdate = workingdate.ToString("yyyyMMdd")
                portcode = embarkport
            ElseIf dayofcruise = maxdays Then
                manifestenddate = workingdate.ToString("yyyyMMdd")
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
        Else
            photogtype = "RPLT"
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

        If manifest = True Then
            'Create the Manifest file
            Dim manifestdt As New DataTable

            manifestdt.Columns.Add("ManifestEntry")


            Dim ManifestPort As String = "MIA"
            Dim ManifestStatus As String = "O"
            Dim manifestSeq As String = "01"
            Dim manifestshipcode As String = Twocharshipcode & "S"
            Dim manifestTower As String = "   "
            Dim ManifestFolder As String = "    "
            Dim manifestGuestFirstName As String = "TEST"
            Dim manifestGuestLastName As String = "GUEST"
            Dim manifestvoyage As String = "0"
      Dim manifestsex As String = "M"
      Dim manifestlang As String = "ENG"
            Dim manifestnationality As String = "USA"
            Dim manifestcabin As String
            Dim manifestFolio As String
            Dim manifestbooking As String
            Dim GroupID As String = "0    "
            Dim groups As String = "0000"
            Dim cabinsection As String = "074"
            Dim Loyalty As String = "G"
            Dim diningtable As String = "0  "
            Dim Diningroom As String = "American Icon - Deck 3   "
            Dim diningtime As String = "MY TIME "
            Dim setsail As String
            Dim guestid As String
            Dim manifestentry As String



      For countermanifest = 0 To numberofguests - 1
        Dim number As Integer
        number = countermanifest + 1
        If Len(validatedvoyage) < 5 Then
          'Add in length padding 0's in the front of the voyage number
          manifestvoyage = validatedvoyage.ToString.PadLeft(5, "0")

        End If

        'Create random cabin, BookingID, GuestID, Folionumber

        Dim foliostart As String = "984100022805"
        Dim foliornd As Integer = RandomNumber(1000, 9999)


        manifestcabin = RandomNumber(1000, 13000).ToString
        manifestbooking = RandomNumber(1000000, 2000000).ToString
        guestid = RandomNumber(100000000, 999999999).ToString

        If runmode = "HOME" Then
          Select Case countermanifest
            Case 0
              manifestFolio = "9841000253346800"
            Case 2
              manifestFolio = "9841000229799801"
            Case 4
              manifestFolio = "9841000139275900"
            Case 6
              manifestFolio = "9841000123604003"
            Case Else
              manifestFolio = foliostart & foliornd.ToString

          End Select
        Else
          manifestFolio = foliostart & foliornd.ToString
        End If

        If Len(manifestcabin) < 5 Then
          manifestcabin = manifestcabin.PadRight(5)
        End If


        'Then create Setsail barcode from GuestID and Cabin
        setsail = guestid & "-" & manifestcabin

        'Create guest name and pad right the string as well.
        Dim guestnumber As String
        Dim manifestguesttmplastname As String
        guestnumber = ConvertNumberToWords(number)
        manifestguesttmplastname = manifestGuestLastName & " " & guestnumber.ToUpper
        manifestGuestFirstName = manifestGuestFirstName.PadRight(15)
        manifestguesttmplastname = manifestguesttmplastname.PadRight(25)

        manifestentry = manifestshipcode & manifestvoyage & manifestStartdate & manifestguesttmplastname & manifestGuestFirstName & ManifestStatus & manifestsex & manifestlang & manifestnationality & manifestcabin & manifestFolio & manifestbooking & manifestSeq & GroupID & groups & cabinsection & Loyalty & manifestStartdate & ManifestPort & manifestenddate & ManifestPort & diningtable & Diningroom & diningtime & manifestTower & ManifestFolder & setsail & guestid
        If DebugON = True Then
          MsgBox(manifestentry)
        End If
        manifestdt.Rows.Add(manifestentry)
      Next

      Dim manifestsb As New StringBuilder
            Dim finalmanifestfile As String = "manifest_" & Twocharshipcode & "_" & validatedvoyage & ".ttx"
            Dim manifestfullpath As String = Path.Combine(userdesktop, finalmanifestfile)

            Dim manifeststreampath As String = manifestfullpath

            'Writes the Stringbuilder to file as a stream. The CSV file is RFC4180 compliant.
            'https://tools.ietf.org/html/rfc4180

            Using writer As StreamWriter = New StreamWriter(manifeststreampath)
                ttxwriter.WriteDataTable(manifestdt, writer, False)
            End Using

        End If
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
    Public Class ttxwriter
        Public Shared Sub WriteDataTable(ByVal sourceTable As DataTable,
            ByVal writer As TextWriter, ByVal includeHeaders As Boolean)
            If (includeHeaders) Then
                Dim headerValues As IEnumerable(Of String) = sourceTable.Columns.OfType(Of DataColumn).Select(Function(column) QuoteValue(column.ColumnName))

                writer.WriteLine(String.Join("", headerValues))
            End If

            Dim items As IEnumerable(Of String) = Nothing
            For Each row As DataRow In sourceTable.Rows
                items = row.ItemArray.Select(Function(obj) QuoteValue(If(obj?.ToString(), String.Empty)))
                writer.WriteLine(String.Join("", items))
            Next

            writer.Flush()
        End Sub

        Private Shared Function QuoteValue(ByVal value As String) As String
            Return String.Concat("""", value.Replace("""", """"""), """")
        End Function
    End Class

    Private Sub DtpEmbarkation_ValueChanged(sender As Object, e As EventArgs) Handles DtpEmbarkation.ValueChanged
        'Checks Date changed and then calculates the end date based on number of days
        Dateadditions()


    End Sub

    Private Sub CboNumberOfDays_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboNumberOfDays.SelectedIndexChanged
        'Checks number of days and then updates the end date.
        Dateadditions()
    End Sub


    Private Sub Dateadditions()
        Dim enddate As DateTime
        If CboNumberOfDays.Text = "0" Then

            Exit Sub
        Else
            If Integer.TryParse(CboNumberOfDays.Text, days) Then
                'Correct
                enddate = DtpEmbarkation.Value.AddDays(days)
                lblEndDatedisplay.Text = enddate.ToString("dddd, MMMM dd, yyyy")
            End If

        End If
    End Sub

    Private Function RandomNumber(min As Integer, max As Integer) As Integer

        Return random.Next(min, max)
    End Function 'RandomNumber




    Private Function GetOnes(ByVal OneDigit As Integer) As String

            GetOnes = ""

            If OneDigit = 0 Then
                Exit Function
            End If

            GetOnes = mOnesArray(OneDigit - 1)

        End Function


        Private Function GetTens(ByVal TensDigit As Integer) As String

            GetTens = ""

            If TensDigit = 0 Or TensDigit = 1 Then
                Exit Function
            End If

            GetTens = mTensArray(TensDigit - 2)

        End Function


        Public Function ConvertNumberToWords(ByVal NumberValue As String) As String

        Dim Delimiter As String = " "
        Dim TensDelimiter As String = "-"
        Dim mNumberValue As String = ""
        Dim mNumbers As String = ""
        Dim mNumWord As String = ""
        Dim mFraction As String = ""
        Dim mNumberStack() As String
        Dim j As Integer = 0
        Dim i As Integer = 0
        Dim mOneTens As Boolean = False

        ConvertNumberToWords = ""

        ' validate input
        Try
            j = CDbl(NumberValue)
        Catch ex As Exception
                ConvertNumberToWords = "Invalid input."
            Exit Function
        End Try

        ' get fractional part {if any}
        If InStr(NumberValue, ".") = 0 Then
            ' no fraction
            mNumberValue = NumberValue
        Else
            mNumberValue = Microsoft.VisualBasic.Left(NumberValue, InStr(NumberValue, ".") - 1)
            mFraction = Mid(NumberValue, InStr(NumberValue, ".")) ' + 1)
            mFraction = Math.Round(CSng(mFraction), 2) * 100

            If CInt(mFraction) = 0 Then
                mFraction = ""
            Else
                mFraction = "&& " & mFraction & "/100"
            End If
        End If
        mNumbers = mNumberValue.ToCharArray

        ' move numbers to stack/array backwards
        For j = mNumbers.Length - 1 To 0 Step -1
            ReDim Preserve mNumberStack(i)

            mNumberStack(i) = mNumbers(j)
            i += 1
        Next

        For j = mNumbers.Length - 1 To 0 Step -1
            Select Case j
                Case 0, 3, 6, 9, 12
                    ' ones  value
                    If Not mOneTens Then
                        mNumWord &= GetOnes(Val(mNumberStack(j))) & Delimiter
                    End If

                    Select Case j
                        Case 3
                            ' thousands
                            mNumWord &= mPlaceValues(1) & Delimiter

                        Case 6
                            ' millions
                            mNumWord &= mPlaceValues(2) & Delimiter

                        Case 9
                            ' billions
                            mNumWord &= mPlaceValues(3) & Delimiter

                        Case 12
                            ' trillions
                            mNumWord &= mPlaceValues(4) & Delimiter
                    End Select


                Case Is = 1, 4, 7, 10, 13
                    ' tens value
                    If Val(mNumberStack(j)) = 0 Then
                        mNumWord &= GetOnes(Val(mNumberStack(j - 1))) & Delimiter
                        mOneTens = True
                        Exit Select
                    End If

                    If Val(mNumberStack(j)) = 1 Then
                        mNumWord &= mOneTensArray(Val(mNumberStack(j - 1))) & Delimiter
                        mOneTens = True
                        Exit Select
                    End If

                    mNumWord &= GetTens(Val(mNumberStack(j)))

                    ' this places the tensdelimiter; check for succeeding 0
                    If Val(mNumberStack(j - 1)) <> 0 Then
                        mNumWord &= TensDelimiter
                    End If
                    mOneTens = False

                Case Else
                    ' hundreds value 
                    mNumWord &= GetOnes(Val(mNumberStack(j))) & Delimiter

                    If Val(mNumberStack(j)) <> 0 Then
                        mNumWord &= mPlaceValues(0) & Delimiter
                    End If
            End Select
        Next

        Return mNumWord & mFraction

    End Function




End Class


