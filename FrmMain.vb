'Copyright (c) 2009-2015 Matt Pittman

#Region "Class / File Comment Header block"

'Program:            Proj03
'File:               Proj03.vb
'Author:             Matt Pittman

'Description:        Part one of the final project. Build a GUI interface that will
'                    eventually support a management system for a theme park. This 
'                    system will house functionality For multiple users
'                    including the customer, sales representative, sales managers
'                    and service operators. The system will be able to store customers,
'                    keep track Of customer's passports and Features that customers
'                    purchase And spend throughout the park. The system will show
'                    summary information and up to date transactions for all passports.

'                    Part two of the final project. Edit GUI as necessary and begin to build 
'                    basic business logic into the project. Add all classes based off UML class
'                    diagram. Add in all attributes needed to all classes. Create all public and
'                    private property procedures for all attributes. Create all public and private ToString
'                    methods. Create special constructor for each class that takes paramaters for
'                    all attributes in the class. Some logic has been put into place such as
'                    user input validation and coutning of objects upon creation.

'                   Part three of the final Project. Gui is polished and we are adding more
'                   logic into the program. We have built custom events for creating of
'                   new objects so that we can keep track of data and pass it around the program.
'                   Calculations now correctly calculate for age and isChild, and methods
'                   can correctly call price based on whether the customer is an adult or child.
'                   All validation is completed with no empty text boxes allowed, no parsing errors,
'                   and no negative numbers in qty fields.
'
'                   Part four of the final project. This part is focused on completing and polishing
'                   all business logic and the final GUI. We have added arrays to store our data,
'                   using iterators to search through our arrays, all calculatins, and reading and writing 
'                   to files. Final touches were made to look clean and professional. Rigorous testing
'                   undergone to be as sure as possible that program will not crash.


'Date:               2015 Sep 14
'                       -Created user interface
'                    2015 Oct 08
'                       -Created Classes
'                       -Created Attributes
'                       -Created Property Procedures and ToString methods
'                       -Created Constructors 
'                       -Built in beginning logic
'                   2015 Nov 5
'                       -Created custom events 
'                       -Filled in data for Process Test Data
'                       -Calculations that we have info for now correctly calculate
'                       -Objects store data in Process Test Data and correctly communicate that Data throughout the method
'                   2015 Dec 01 
'                       -Arrays Added
'                       -Searching usings ID's now works
'                   2015 Dec 06
'                       -Read/Write to File working
'                   2015 Dec 08 
'                       -Metrics added 


'Tier:               User Interface
'Exceptions:         No exceptions are defined
'Exception-Handling: No exceptions are handled

'Events:            Create Feature
'                   Create Customer
'                   Create Passport
'                   Create PassportFeature
'                   Use Feature
'                   Update Passport Feature
'                   

'Event-Handling:     Buttons all call methods to create objects

#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
Imports System.IO
#End Region 'Option / Imports

Public Class FrmMain


#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables

    Private WithEvents mThemePark As ThemePark

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.
    'These are all public.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _theThemePark As ThemePark
        Get
            Return mThemePark
        End Get
        Set(pValue As ThemePark)
            mThemePark = pValue
        End Set
    End Property

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    'No Behavioral Methods are currently defined.

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    '********** Private Non-Shared Behavioral Methods

    Private Sub _initializeBusinessLogic()
        '_initializeBusinessLogic() will add any business logic prior to any user input

        _theThemePark = New ThemePark("CIS605 Theme Park")
        Me.Text = "CIS605 Theme Park"


    End Sub '_initializeBusinessLogic()

    Private Sub _initializeUserInterface()
        '_initializeUserInterface() will make any necessary changes to the User Interface before the user views it

        Me.AcceptButton = btnProcessTestDataTabTransactionsTbcMain
        Me.CancelButton = btnExit
        txtCustomerNameTabCustomerTbcMain.Focus()
        '_processTestData()
        ' _theThemePark.readFile()

    End Sub '_initializeUserInterface()


    Private Sub _processTestData()

        Dim date1 As New Date(2015, 12, 12)
        Dim qty As Integer = 5


        Dim c01, c02, c03 As Customer
        lstCustomerListTabSummaryInfoTbcMain.Items.Add("Customer")
        lstCustomerListTabSummaryInfoTbcMain.Items.Add("ID")
        lstCustomerListTabSummaryInfoTbcMain.Items.Add("Here")
        cboCustomerNameTabPassportsTbcMain.Items.Add("Customer")
        cboCustomerNameTabPassportsTbcMain.Items.Add("ID")
        cboCustomerNameTabPassportsTbcMain.Items.Add("Here")
        c01 = _theThemePark.addCustomer("C01", "CName01")
        c02 = _theThemePark.addCustomer("C02", "CName02")
        c03 = _theThemePark.addCustomer("C03", "Customer Name 03")


        txtTransactions.Text &= vbCrLf & "******************PROCESS TEST DATA IS CREATING ONE OBJECT FOR EACH CLASS******************" & vbCrLf


        Dim f01, f02, f03 As Feature
        lstFeatureListTabSummaryInfoTbcMain.Items.Add("Feature")
        lstFeatureListTabSummaryInfoTbcMain.Items.Add("ID")
        lstFeatureListTabSummaryInfoTbcMain.Items.Add("Here")
        cboFeatureIdTabAddTabPassportFeaturesTbcMain.Items.Add("Feature")
        cboFeatureIdTabAddTabPassportFeaturesTbcMain.Items.Add("ID")
        cboFeatureIdTabAddTabPassportFeaturesTbcMain.Items.Add("Here")
        f01 = _theThemePark.addFeature("F01", "Park Pass", "Day", 100, 80)
        f02 = _theThemePark.addFeature("F02", "Early Entry Pass", "Day", 10, 5)
        f03 = _theThemePark.addFeature("F03", "Meal Plan", "Meal", 30, 20)


        Dim pb01, pb02, pb03, pb04, pb05, pb06 As Passport
        lstPassportListTabSummaryInfoTbcMain.Items.Add("Passport")
        lstPassportListTabSummaryInfoTbcMain.Items.Add("ID")
        lstPassportListTabSummaryInfoTbcMain.Items.Add("Here")
        cboPassportIdTabAddTabPassportFeaturesTbcMain.Items.Add("Passport")
        cboPassportIdTabAddTabPassportFeaturesTbcMain.Items.Add("ID")
        cboPassportIdTabAddTabPassportFeaturesTbcMain.Items.Add("Here")
        cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Items.Add("Passport")
        cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Items.Add("ID")
        cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Items.Add("Here")
        cboPassportIdTabUseFeatureTbcMain.Items.Add("Passport")
        cboPassportIdTabUseFeatureTbcMain.Items.Add("ID")
        cboPassportIdTabUseFeatureTbcMain.Items.Add("Here")
        pb01 = _theThemePark.addPassport("PB01", New Date(2015, 9, 15), "self", New Date(1980, 1, 1), c01)
        pb02 = _theThemePark.addPassport("PB02", New Date(2015, 9, 16), "self", New Date(1985, 6, 1), c02)
        pb03 = _theThemePark.addPassport("PB03", New Date(2015, 9, 17), "C02 Visitor", New Date(2003, 12, 1), c02)
        pb04 = _theThemePark.addPassport("PB04", New Date(2015, 8, 15), "self", New Date(1975, 1, 1), c03)
        pb05 = _theThemePark.addPassport("PB05", New Date(2015, 9, 15), "C03 Visitor 1", New Date(2002, 10, 7), c03)
        pb06 = _theThemePark.addPassport("PB06", New Date(2015, 10, 15), "C03 Visitor 2", New Date(2002, 10, 8), c03)



        Dim pbf01, pbf02, pbf03, pbf04, pbf05, pbf06, pbf07, pbf08, pbf09, pbf10 As PassportFeature
        lstPassportFeatureListTabSummaryInfoTbcMain.Items.Add("Passport")
        lstPassportFeatureListTabSummaryInfoTbcMain.Items.Add("ID")
        lstPassportFeatureListTabSummaryInfoTbcMain.Items.Add("Here")
        pbf01 = _theThemePark.addPassportFeature("PBF01", 1, pb01.returnPrice(pb01, f01, (pb01.isChildUnder13(pb01.calcAge(pb01.Birthdate, pb01.DatePurchased)))), pb01, f01)
        pbf02 = _theThemePark.addPassportFeature("PBF02", 2, pb02.returnPrice(pb02, f01, (pb02.isChildUnder13(pb02.calcAge(pb02.Birthdate, pb02.DatePurchased)))), pb02, f01)
        pbf03 = _theThemePark.addPassportFeature("PBF03", 3, pb03.returnPrice(pb03, f01, (pb03.isChildUnder13(pb03.calcAge(pb03.Birthdate, pb03.DatePurchased)))), pb03, f01)
        pbf04 = _theThemePark.addPassportFeature("PBF04", 1, pb04.returnPrice(pb04, f01, (pb04.isChildUnder13(pb04.calcAge(pb04.Birthdate, pb04.DatePurchased)))), pb04, f01)
        pbf05 = _theThemePark.addPassportFeature("PBF05", 1, pb05.returnPrice(pb05, f01, (pb05.isChildUnder13(pb05.calcAge(pb05.Birthdate, pb05.DatePurchased)))), pb05, f01)
        pbf06 = _theThemePark.addPassportFeature("PBF06", 1, pb06.returnPrice(pb06, f01, (pb06.isChildUnder13(pb06.calcAge(pb06.Birthdate, pb06.DatePurchased)))), pb06, f01)
        pbf07 = _theThemePark.addPassportFeature("PBF07", 3, pb03.returnPrice(pb03, f02, (pb03.isChildUnder13(pb03.calcAge(pb03.Birthdate, pb03.DatePurchased)))), pb03, f02)
        pbf08 = _theThemePark.addPassportFeature("PBF08", 9, pb03.returnPrice(pb03, f03, (pb03.isChildUnder13(pb03.calcAge(pb03.Birthdate, pb03.DatePurchased)))), pb03, f03)
        pbf09 = _theThemePark.addPassportFeature("PBF09", 1, pb04.returnPrice(pb04, f01, (pb01.isChildUnder13(pb01.calcAge(pb01.Birthdate, pb01.DatePurchased)))), pb04, f01)
        pbf10 = _theThemePark.addPassportFeature("PBF10", 3, pb04.returnPrice(pb04, f01, (pb04.isChildUnder13(pb04.calcAge(pb04.Birthdate, pb04.DatePurchased)))), pb04, f01)



        Dim uf01, uf02, uf03, uf04 As UsedFeature
        lstUsedFeatureListTabSummaryInfoTbcMain.Items.Add("Passport")
        lstUsedFeatureListTabSummaryInfoTbcMain.Items.Add("ID")
        lstUsedFeatureListTabSummaryInfoTbcMain.Items.Add("Here")
        uf01 = _theThemePark.addUsedFeature("UF01", New Date(2015, 10, 20), "Epcot Center", 1, pbf01)
        uf02 = _theThemePark.addUsedFeature("UF02", New Date(2015, 10, 20), "West Parking", 1, pbf02)
        uf03 = _theThemePark.addUsedFeature("UF03", New Date(2015, 10, 20), "France", 2, pbf03)
        uf04 = _theThemePark.addUsedFeature("UF04", New Date(2015, 10, 20), "American Pavilion", 1, pbf03)



        _theThemePark.updatePassportFeature("PBF03", 1, pbf03.Price, pbf03, pbf03.Passport, pbf03.Feature, Date.Today)



        txtTransactions.Text &= vbCrLf & _theThemePark.ToString()

        txtTransactions.Text &= vbCrLf & "******************PROCESS TEST DATA HAS FINISHED CREATING ONE OBJECT FOR EACH CLASS******************" & vbCrLf & vbCrLf

        btnProcessTestDataTabTransactionsTbcMain.Enabled = False

    End Sub

    Private Sub _refreshMetrics()


        txtMetricsTabSummaryInfoTbcMain.Text = ""

        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.averageBalanceUnused()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.sumUnusedPassportFeature()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.averagePassportsPerCustomer()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.mostPopularPassportFeature()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.percentPassportFeaturesUsed()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.averageAge()
        txtMetricsTabSummaryInfoTbcMain.Text &= _theThemePark.passportHoldersBirthday()


    End Sub


#End Region 'Behavioral Methods


#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined.
    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user


    Private Sub btnAddCustomer_Click(sender As Object, e As EventArgs) Handles btnAddCustomer.Click

        If txtCustomerNameTabCustomerTbcMain.Text.Trim.Length = 0 Then
            MessageBox.Show("Please enter customer Name")
            txtCustomerNameTabCustomerTbcMain.Select()
            txtCustomerNameTabCustomerTbcMain.Focus()
            Exit Sub
        End If

        If txtCustomerIdTabCustomerTbcMain.Text.Trim.Length = 0 Then
            MessageBox.Show("Please enter customer ID")
            txtCustomerIdTabCustomerTbcMain.Select()
            txtCustomerIdTabCustomerTbcMain.Focus()
            Exit Sub
        End If


        Dim inputValue As String
        Dim foundValue As Customer
        Dim foundLocation As Integer

        'get/validate input

        inputValue = txtCustomerIdTabCustomerTbcMain.Text

        'do processing

        foundValue = _theThemePark.findCustomer(inputValue, foundLocation)

        'display info

        If foundValue Is Nothing Then 'NOT found

            Dim customer1 = _theThemePark.addCustomer(
            txtCustomerIdTabCustomerTbcMain.Text,
            txtCustomerNameTabCustomerTbcMain.Text
            )

            txtCustomerIdTabCustomerTbcMain.Clear()
            txtCustomerNameTabCustomerTbcMain.Clear()
            txtCustomerNameTabCustomerTbcMain.SelectAll()
            txtCustomerNameTabCustomerTbcMain.Focus()

        Else 'FOUND

            MessageBox.Show("Customer ID '" & inputValue & "' already exists" & vbCrLf _
                        & "Please enter a customer ID not already in use.")

            txtCustomerIdTabCustomerTbcMain.SelectAll()
            txtCustomerIdTabCustomerTbcMain.Focus()

            'End If

        End If





    End Sub 'btnAddCustomer_Click (sender,e)

    Private Sub btnAddFeatureTabFeaturesTbcMain_Click(sender As Object, e As EventArgs) Handles btnAddFeatureTabFeaturesTbcMain.Click

        Dim adultPricePerUnit As Decimal
        Dim childPricePerUnit As Decimal

        If txtFeatureIdTabFeaturesTbcMain.Text.Trim.Length = 0 Then
            MessageBox.Show("Please enter Featre ID")
            txtFeatureIdTabFeaturesTbcMain.Select()
            txtFeatureIdTabFeaturesTbcMain.Focus()
            Exit Sub
        End If

        If txtFeatureNameTabFeaturesTbcMain.Text.Trim.Length = 0 Then
            MessageBox.Show("Please enter Feature Name")
            txtFeatureNameTabFeaturesTbcMain.Select()
            txtFeatureNameTabFeaturesTbcMain.Focus()
            Exit Sub
        End If

        If txtUnitOfMeasureTabFeaturesTbcMain.Text.Trim.Length = 0 Then
            MessageBox.Show("Please enter Unit of Measure")
            txtUnitOfMeasureTabFeaturesTbcMain.Select()
            txtUnitOfMeasureTabFeaturesTbcMain.Focus()
            Exit Sub
        End If

        Try
            adultPricePerUnit = Decimal.Parse(txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Text)
        Catch ex As Exception
            MessageBox.Show(
            "Error: Invalid Price. " _
            & "Please enter a number for the Price.  " _
            & "Ex: 20 or 25.50"
            )
            txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.SelectAll()
            txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Focus()
            Exit Sub
        End Try
        adultPricePerUnit = Decimal.Parse(txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Text)

        If adultPricePerUnit < 0 Then
            MessageBox.Show("Please enter positive Price")
            txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Select()
            txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Focus()
            Exit Sub
        End If

        Try
            childPricePerUnit = Decimal.Parse(txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Text)
        Catch ex As Exception
            MessageBox.Show(
            "ERROR: Invalid Price. " _
            & "Please enter a number for the Price.  " _
            & "Ex: 20 or 25.50"
            )
            txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.SelectAll()
            txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Focus()
            Exit Sub
        End Try
        childPricePerUnit = Decimal.Parse(txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Text)

        If childPricePerUnit < 0 Then
            MessageBox.Show("Please enter positive Price")
            txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Select()
            txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Focus()
            Exit Sub
        End If

        Dim inputValue As String
        Dim foundValue As Feature
        Dim foundLocation As Integer

        'get/validate input

        inputValue = txtFeatureIdTabFeaturesTbcMain.Text

        'do processing

        foundValue = _theThemePark.findFeature(inputValue, foundLocation)

        'display info

        If foundValue Is Nothing Then 'NOT found

            Dim feature1 = _theThemePark.addFeature(
            txtFeatureIdTabFeaturesTbcMain.Text,
            txtFeatureNameTabFeaturesTbcMain.Text,
            txtUnitOfMeasureTabFeaturesTbcMain.Text,
            adultPricePerUnit,
            childPricePerUnit
            )

            txtFeatureIdTabFeaturesTbcMain.Clear()
            txtFeatureNameTabFeaturesTbcMain.Clear()
            txtUnitOfMeasureTabFeaturesTbcMain.Clear()
            txtAdultPriceGrpFeaturePriceTabFeaturesTbcMain.Clear()
            txtChildPriceGrpFeaturePriceTabFeaturesTbcMain.Clear()
            txtFeatureIdTabFeaturesTbcMain.Select()
            txtFeatureIdTabFeaturesTbcMain.Focus()

        Else 'FOUND

            MessageBox.Show("Feature ID '" & inputValue & "' already exists" & vbCrLf _
                        & "Please enter a Feature ID not already in use.")

            txtFeatureIdTabFeaturesTbcMain.SelectAll()
            txtFeatureIdTabFeaturesTbcMain.Focus()

            'End If

        End If



    End Sub 'btnAddFeatureTabFeaturesTbcMain_Click

    Private Sub btnAddPassportGrpVisitorInformationTabPassportsTbcMain_Click(sender As Object, e As EventArgs) _
        Handles btnAddPassportGrpVisitorInformationTabPassportsTbcMain.Click

        Dim visitorBirthdate, datePurchased As Date
        Dim age As Integer
        Dim isChild As Boolean
        Dim theCustomer As Customer
        Dim foundLocation As Integer

        visitorBirthdate = dtpVisitorBirthdateGrpVisitorInformationTabPassportsTbcMain.Value
        datePurchased = Date.Today
        theCustomer = _theThemePark.findCustomer(cboCustomerNameTabPassportsTbcMain.Text, foundLocation)

        If theCustomer Is Nothing Then
            MessageBox.Show("Customer does not exist, please select a valid customer")
            cboCustomerNameTabPassportsTbcMain.SelectAll()
            cboCustomerNameTabPassportsTbcMain.Focus()
            Exit Sub
        Else



            If cboCustomerNameTabPassportsTbcMain.SelectedIndex = -1 Then
                MessageBox.Show("Please Select a Customer ID from the drop down menu")
                cboCustomerNameTabPassportsTbcMain.Select()
                cboCustomerNameTabPassportsTbcMain.Focus()
                Exit Sub
            End If


            If txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter Passport ID")
                txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Select()
                txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Focus()
                Exit Sub
            End If

            If txtVisitorNameGrpVisitorInformationTabPassportsTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter Visitor Name")
                txtVisitorNameGrpVisitorInformationTabPassportsTbcMain.Select()
                txtVisitorNameGrpVisitorInformationTabPassportsTbcMain.Focus()
                Exit Sub
            End If

            Dim inputValue As String
            Dim foundValue As Passport
            Dim foundPassportLocation As Integer

            'get/validate input

            inputValue = txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Text

            'do processing

            foundValue = _theThemePark.findPassport(inputValue, foundPassportLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found

                Dim thePassport = _theThemePark.addPassport(
            txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Text,
            datePurchased,
            txtVisitorNameGrpVisitorInformationTabPassportsTbcMain.Text,
            visitorBirthdate,
            theCustomer
            )

                txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Clear()
                txtVisitorNameGrpVisitorInformationTabPassportsTbcMain.Clear()
                cboCustomerNameTabPassportsTbcMain.SelectedIndex = -1
                dtpVisitorBirthdateGrpVisitorInformationTabPassportsTbcMain.Value = Today
                cboCustomerNameTabPassportsTbcMain.Select()
                cboCustomerNameTabPassportsTbcMain.Focus()

            Else 'FOUND

                MessageBox.Show("Passport ID '" & inputValue & "' already exists" & vbCrLf _
                        & "Please enter a Passport ID not already in use.")

                txtPassortIdGrpVisitorInformationTabPassportsTbcMain.SelectAll()
                txtPassortIdGrpVisitorInformationTabPassportsTbcMain.Focus()

                'End If

            End If
        End If




    End Sub 'btnAddPassportGrpVisitorInformationTabPassportsTbcMain_Click


    Private Sub btnAddPassportFeatureGrpPriceQuantityTabAddTabPassportFeaturesTbcMain_Click(sender As Object, e As EventArgs) Handles btnAddPassportFeatureGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Click


        ' Dim date1 As New Date(2015, 12, 12)
        Dim theFeature As Feature
        Dim theCustomer As Customer
        Dim thePassport As Passport
        Dim price, quantity As Decimal
        Dim foundLocation As Integer




        thePassport = _theThemePark.findPassport(cboPassportIdTabAddTabPassportFeaturesTbcMain.Text, foundLocation)


        If thePassport Is Nothing Then
            MessageBox.Show("Passport does not exist, please select a valid passport")
            cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectAll()
            cboPassportIdTabAddTabPassportFeaturesTbcMain.Focus()
            Exit Sub
        Else

            theCustomer = thePassport.Owner
            theFeature = _theThemePark.findFeature(cboFeatureIdTabAddTabPassportFeaturesTbcMain.Text, foundLocation)

            If theFeature Is Nothing Then
                MessageBox.Show("Feature does not exist, please select a valid feature")
                cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectAll()
                cboFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()
                Exit Sub
            Else

                Dim isChildUnder13 As Boolean
                Dim age As Integer
                isChildUnder13 = False
                age = thePassport.calcAge(thePassport.Birthdate, Today)
                isChildUnder13 = thePassport.isChildUnder13(age)
                price = thePassport.returnPrice(thePassport, theFeature, isChildUnder13)


                If txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Text.Trim.Length = 0 Then
                    MessageBox.Show("Please enter Passport Feature ID")
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Select()
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End If


                If cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = -1 Then
                    MessageBox.Show("Please Select a Passport ID from the drop down menu")
                    cboPassportIdTabAddTabPassportFeaturesTbcMain.Select()
                    cboPassportIdTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End If

                If cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = -1 Then
                    MessageBox.Show("Please Select a Feature ID from the drop down menu")
                    cboFeatureIdTabAddTabPassportFeaturesTbcMain.Select()
                    cboFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End If

                'Dont think it is necessary to have an input for ID when the user already selects it, leaving here just in case
                'If txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Text.Trim.Length = 0 Then
                '    MessageBox.Show("Please enter Passport Feature ID")
                '    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Select()
                '    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()
                '    Exit Sub
                'End If


                If txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text.Trim.Length = 0 Then
                    MessageBox.Show("Please enter Quantity")
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Select()
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End If



                Try
                    quantity = Decimal.Parse(txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text)
                Catch ex As Exception
                    MessageBox.Show(
                    "ERROR: Invalid Quantity Remaining. " _
                    & "Please enter a quantity in numbers.  " _
                    & "Ex: 7 or 22"
                    )
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.SelectAll()
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End Try
                quantity = Decimal.Parse(txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text)

                If quantity < 0 Then
                    MessageBox.Show("Please enter positive Quantity")
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Select()
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Focus()
                    Exit Sub
                End If

                Dim inputValue As String
                Dim foundValue As PassportFeature
                Dim foundpassportFeatureLocation As Integer

                'get/validate input

                inputValue = txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Text

                'do processing

                foundValue = _theThemePark.findPassportFeature(inputValue, foundpassportFeatureLocation)

                'display info

                If foundValue Is Nothing Then 'NOT found

                    Dim passportFeature1 = _theThemePark.addPassportFeature(
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Text,
                    quantity,
                    price,
                    thePassport,
                    theFeature
                    )

                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Clear()
                    txtAdultPriceGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Clear()
                    txtQuantityGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Clear()
                    cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = -1
                    cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = -1
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Select()
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()


                Else 'FOUND

                    MessageBox.Show("PassportFeature ID '" & inputValue & "' already exists" & vbCrLf _
                                & "Please enter a PassportFeature ID not already in use.")

                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.SelectAll()
                    txtPassportFeatureIdTabAddTabPassportFeaturesTbcMain.Focus()

                    'End If

                End If

            End If

        End If

    End Sub 'btnAddPassportFeatureGrpPriceQuantityTabAddTabPassportFeaturesTbcMain_Click


    Private Sub btnUpdatePassportFeatureGrpPriceQuantityTabAddTabPassportFeaturesTbcMain_Click(sender As Object, e As EventArgs) Handles btnUpdatePassportFeatureGrpAddTabUpdateTabPassportFeaturesTbcMain.Click
        '
        'Notes saying to --- Looking ahead on Update PassbookFeature -- in project 4,
        'instead of passing in the original passbook feature as a
        'parameter, you'll be better off by passing in the ID and then
        ' finding the correct object in the array.
        '

        'Dim date1 As New Date(2015, 12, 12)
        Dim theFeature As Feature
        Dim theCustomer As Customer
        Dim thePassport As Passport
        Dim thePassportFeature As PassportFeature
        Dim price, quantity As Decimal
        Dim foundLocation As Integer
        Dim isChildUnder13 As Boolean
        Dim age As Integer

        thePassportFeature = _theThemePark.findPassportFeature(cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Text, foundLocation)

        If thePassportFeature Is Nothing Then
            MessageBox.Show("Passport Feature does not exist, please select a valid passport feature")
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectAll()
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Focus()
            Exit Sub
        Else

            thePassport = thePassportFeature.Passport
            theFeature = thePassportFeature.Feature
            theCustomer = thePassport.Owner
            age = thePassport.calcAge(thePassport.Birthdate, Today)
            isChildUnder13 = thePassport.isChildUnder13(age)
            price = thePassport.returnPrice(thePassport, theFeature, isChildUnder13)

            price = thePassport.returnPrice(thePassport, theFeature, isChildUnder13)


            If cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedIndex = -1 Then
                MessageBox.Show("Please Select a Passport ID from the drop down menu")
                cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Select()
                cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Focus()
                Exit Sub
            End If


            'If txtPassportFeatureIdGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text.Trim.Length = 0 Then
            '    MessageBox.Show("Please enter Passport Feature ID")
            '    txtPassportFeatureIdGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Select()
            '    txtPassportFeatureIdGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Focus()
            '    Exit Sub
            'End If


            If txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter Quantity")
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Select()
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Focus()
                Exit Sub
            End If


            Try
                quantity = Decimal.Parse(txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text)
            Catch ex As Exception
                MessageBox.Show(
            "ERROR: Invalid Quantity Remaining. " _
            & "Please enter a quantity in numbers.  " _
            & "Ex: 7 or 22"
            )
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.SelectAll()
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Focus()
                Exit Sub
            End Try
            quantity = Decimal.Parse(txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text)

            If quantity < 0 Then
                MessageBox.Show("Please enter positive Quantity")
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Select()
                txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Focus()
                Exit Sub
            End If



            Dim aPassport = _theThemePark.updatePassportFeature(
        cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Text,
        quantity,
        price,
        thePassportFeature,
        thePassport,
        theFeature,
        Date.Today
      )

            'txtPassportFeatureIdGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Clear()
            txtPriceGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Clear()
            txtQtyGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Clear()
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedIndex = -1
            txtFeatureToStringTabUpdateTabPassportFeaturesTbcMain.Clear()
            txtCurrentQuantityGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Clear()
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Select()
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Focus()
        End If


    End Sub 'btnAddPassportFeatureGrpPriceQuantityTabAddTabPassportFeaturesTbcMain_Click

    Private Sub btnUseFeatureGrpFeatureDetailsTabUseFeatureTbcMain_Click(sender As Object, e As EventArgs) Handles btnUseFeatureGrpFeatureDetailsTabUseFeatureTbcMain.Click


        Dim quantityUsed As Decimal
        Dim theCustomer As Customer
        Dim thePassport As Passport
        Dim thePassportFeature As PassportFeature
        Dim price, quantity As Decimal
        Dim foundLocation As Integer
        Dim isChildUnder13 As Boolean
        Dim age As Integer


        thePassportFeature = _theThemePark.findPassportFeature(cboPassportIdTabUseFeatureTbcMain.Text, foundLocation)

        If thePassportFeature Is Nothing Then
            MessageBox.Show("Passport Feature does not exist, please select a passport feature")
            cboPassportIdTabUseFeatureTbcMain.SelectAll()
            cboPassportIdTabUseFeatureTbcMain.Focus()
            Exit Sub
        Else


            thePassport = thePassportFeature.Passport
            theCustomer = thePassport.Owner
            age = thePassport.calcAge(thePassport.Birthdate, Today)
            isChildUnder13 = thePassport.isChildUnder13(age)
            price = thePassport.returnPrice(thePassport, thePassportFeature.Feature, isChildUnder13)


            If txtEntitlementIdTabUseFeatureTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter Entitlement ID")
                txtEntitlementIdTabUseFeatureTbcMain.Select()
                txtEntitlementIdTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            If cboPassportIdTabUseFeatureTbcMain.SelectedIndex = -1 Then
                MessageBox.Show("Please Select a Passport ID from the drop down menu")
                cboPassportIdTabUseFeatureTbcMain.Select()
                cboPassportIdTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            'If cboFeatureIdTabUseFeatureTbcMain.SelectedIndex = -1 Then
            '    MessageBox.Show("Please Select a Feature ID from the drop down menu")
            '    cboFeatureIdTabUseFeatureTbcMain.Select()
            '    cboFeatureIdTabUseFeatureTbcMain.Focus()
            '    Exit Sub
            'End If

            If txtLocationGrpFeatureDetailsTabUseFeatureTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter a location")
                txtLocationGrpFeatureDetailsTabUseFeatureTbcMain.Select()
                txtLocationGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            If txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Text.Trim.Length = 0 Then
                MessageBox.Show("Please enter a quantity")
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Select()
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            Try
                quantityUsed = Decimal.Parse(txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Text)
            Catch ex As Exception
                MessageBox.Show(
            "ERROR: Invalid Age. " _
            & "Please enter a age in years.  " _
            & "Ex: 7 or 22"
            )
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.SelectAll()
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                Exit Sub
            End Try
            quantityUsed = Decimal.Parse(txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Text)

            If quantityUsed < 0 Then
                MessageBox.Show("Please enter a positive quantity")
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Select()
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            If quantityUsed = 0 Then
                MessageBox.Show("Please enter a positive number above 0")
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Select()
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                Exit Sub
            End If

            'Check to see if ID already exist
            Dim inputValue As String
            Dim foundValue As UsedFeature
            Dim foundUsedFeatureLocation As Integer

            'get/validate input

            inputValue = txtEntitlementIdTabUseFeatureTbcMain.Text

            'do processing

            foundValue = _theThemePark.findUsedFeature(inputValue, foundUsedFeatureLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found

                If quantityUsed > thePassportFeature.Quantity Then

                    MessageBox.Show("Quantity: '" & quantityUsed & "' exceeds the current quantity of: '" & thePassportFeature.Quantity & "', Please re-enter a lower quantity.")
                    txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.SelectAll()
                    txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Focus()
                    Exit Sub

                End If

                Dim aUsedFeature = _theThemePark.addUsedFeature(
            txtEntitlementIdTabUseFeatureTbcMain.Text,
            Date.Today,
            txtLocationGrpFeatureDetailsTabUseFeatureTbcMain.Text,
            quantityUsed,
            thePassportFeature
            )


                txtEntitlementIdTabUseFeatureTbcMain.Clear()
                txtLocationGrpFeatureDetailsTabUseFeatureTbcMain.Clear()
                txtQuantityUsedGrpFeatureDetailsTabUseFeatureTbcMain.Clear()
                txtCustomerToStringTabUseFeatureTbcMain.Clear()
                txtVisitorToStringTabUseFeatureTbcMain.Clear()
                txtFeatureToStringTabUseFeatureTbcMain.Clear()
                txtRemainingQuantityGrpFeatureDetailsTabUseFeatureTbcMain.Clear()
                txtPreviouslyUsedTabUseFeatureTbcMain.Clear()
                cboPassportIdTabUseFeatureTbcMain.SelectedIndex = -1
                txtEntitlementIdTabUseFeatureTbcMain.Select()
                txtEntitlementIdTabUseFeatureTbcMain.Focus()


            Else 'FOUND

                MessageBox.Show("Entitlement ID '" & inputValue & "' already exists" & vbCrLf _
                        & "Please enter a Entitlement ID not already in use.")

                txtEntitlementIdTabUseFeatureTbcMain.SelectAll()
                txtEntitlementIdTabUseFeatureTbcMain.Focus()

                'End If

            End If
        End If




    End Sub 'btnUseFeatureGrpFeatureDetailsTabUseFeatureTbcMain_Click

    Private Sub btnReadFileTabTransactionsTbcMain_Click(sender As Object, e As EventArgs) Handles btnReadFileTabTransactionsTbcMain.Click

        _theThemePark.readFile()

    End Sub



    Private Sub btnWriteFileTabTransactionsTbcMain_Click(sender As Object, e As EventArgs) Handles btnWriteFileTabTransactionsTbcMain.Click



        If chkAppendTabTransactionsTbcMain.Checked Then
            _theThemePark.writeFile(True)
        Else
            _theThemePark.writeFile(False)
        End If


    End Sub


    Private Sub btnProcessTestDataTabTransactionsTbcMain_Click(sender As Object, e As EventArgs) Handles btnProcessTestDataTabTransactionsTbcMain.Click

        _processTestData()

    End Sub 'btnProcessTestDataTabTransactionsTbcMain_Click(sender,e)

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MessageBox.Show("CIS605 Theme Park Management System a.k.a 'Super Phun Thyme!'" & vbCrLf & "Devolped by Matt Pittman")


    End Sub





    Private Sub _btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub '_btnExit_Click(sender,e)




    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        _initializeBusinessLogic()
        _initializeUserInterface()



    End Sub 'FrmMain_Load

    Private Sub _txtTrx_TextChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            txtTransactions.TextChanged

        txtTransactions.SelectionStart = txtTransactions.TextLength
        txtTransactions.ScrollToCaret()

    End Sub '_txtTrx_TextChanged(sender,e)

    Private Sub _lstCustomerListTabSummaryInfoTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles lstCustomerListTabSummaryInfoTbcMain.SelectedIndexChanged

        txtToStringTabSummaryInfoTbcMain.Text = ""


        If lstCustomerListTabSummaryInfoTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = lstCustomerListTabSummaryInfoTbcMain.SelectedIndex

            Dim inputValue As String
            Dim foundValue As Customer
            Dim foundLocation As Integer

            'get/validate input

            inputValue = lstCustomerListTabSummaryInfoTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findCustomer(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtToStringTabSummaryInfoTbcMain.Text &=
                    "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtToStringTabSummaryInfoTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                'End If

            End If
        End If
        'get ready for next input

    End Sub '_lstCustomerListTabSummaryInfoTbcMain_SelectedIndexChanged

    Private Sub _lstFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles lstFeatureListTabSummaryInfoTbcMain.SelectedIndexChanged

        txtToStringTabSummaryInfoTbcMain.Text = ""


        If lstFeatureListTabSummaryInfoTbcMain.SelectedIndex >= 0 Then

            Dim selectedIndex As Integer = lstFeatureListTabSummaryInfoTbcMain.SelectedIndex
            Dim inputValue As String
            Dim foundValue As Feature
            Dim foundLocation As Integer

            'get/validate input

            inputValue = lstFeatureListTabSummaryInfoTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findFeature(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtToStringTabSummaryInfoTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtToStringTabSummaryInfoTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                'End If

            End If
        End If

    End Sub '_lstFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged

    Private Sub _lstPassportFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles lstPassportFeatureListTabSummaryInfoTbcMain.SelectedIndexChanged

        txtToStringTabSummaryInfoTbcMain.Text = ""


        If lstPassportFeatureListTabSummaryInfoTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = lstPassportFeatureListTabSummaryInfoTbcMain.SelectedIndex


            Dim inputValue As String
            Dim foundValue As PassportFeature
            Dim foundLocation As Integer

            'get/validate input

            inputValue = lstPassportFeatureListTabSummaryInfoTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findPassportFeature(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtToStringTabSummaryInfoTbcMain.Text &=
                 "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtToStringTabSummaryInfoTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                'End If

            End If

        End If

    End Sub '_lstPassportFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged

    Private Sub _lstPassportListTabSummaryInfoTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles lstPassportListTabSummaryInfoTbcMain.SelectedIndexChanged

        txtToStringTabSummaryInfoTbcMain.Text = ""


        If lstPassportListTabSummaryInfoTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = lstPassportListTabSummaryInfoTbcMain.SelectedIndex


            Dim inputValue As String
            Dim foundValue As Passport
            Dim foundLocation As Integer

            'get/validate input

            inputValue = lstPassportListTabSummaryInfoTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findPassport(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtToStringTabSummaryInfoTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtToStringTabSummaryInfoTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                'End If

            End If

        End If

    End Sub '_lstPassportListTabSummaryInfoTbcMain_SelectedIndexChanged

    Private Sub _lstUsedFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles lstUsedFeatureListTabSummaryInfoTbcMain.SelectedIndexChanged

        txtToStringTabSummaryInfoTbcMain.Text = ""


        If lstUsedFeatureListTabSummaryInfoTbcMain.SelectedIndex >= 0 Then

            Dim selectedIndex As Integer = lstUsedFeatureListTabSummaryInfoTbcMain.SelectedIndex


            Dim inputValue As String
            Dim foundValue As UsedFeature
            Dim foundLocation As Integer


            inputValue = lstUsedFeatureListTabSummaryInfoTbcMain.SelectedItem.ToString
            foundValue = _theThemePark.findUsedFeature(inputValue, foundLocation)


            If foundValue Is Nothing Then 'NOT found
                txtToStringTabSummaryInfoTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtToStringTabSummaryInfoTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                'End If

            End If

        End If




    End Sub '_lstUsedFeatureListTabSummaryInfoTbcMain_SelectedIndexChanged

    Private Sub _cboCustomerNameTabPassportsTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles cboCustomerNameTabPassportsTbcMain.SelectedIndexChanged

        txtCustomerToStringTabPassportsTbcMain.Text = ""

        If cboCustomerNameTabPassportsTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = cboCustomerNameTabPassportsTbcMain.SelectedIndex

            Dim inputValue As String
            Dim foundValue As Customer
            Dim foundLocation As Integer

            'get/validate input

            inputValue = cboCustomerNameTabPassportsTbcMain.Text

            'do processing

            foundValue = _theThemePark.findCustomer(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtCustomerToStringTabPassportsTbcMain.Text &=
                  "'" & inputValue & "' NOT found."
            Else 'FOUND


                txtCustomerToStringTabPassportsTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf


                'End If
                txtTransactions.Text &=
                    vbCrLf
            End If
        End If

    End Sub '_cboCustomerNameTabPassportsTbcMain_SelectedIndexChanged


    Private Sub _cboPassportIdTabAddTabPassportFeaturesTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndexChanged


        txtCustomerToStringTabAddTabPassportFeaturesTbcMain.Text = ""
        txtVisitorToStringTabAddTabPassportFeaturesTbcMain.Text = ""

        If cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex
            Dim featureSelectedIndex As Integer = cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex

            Dim inputValue As String
            Dim foundValue As Passport
            Dim foundLocation As Integer

            'get/validate input

            inputValue = cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findPassport(inputValue, foundLocation)

            'display info

            If foundValue Is Nothing Then 'NOT found
                txtCustomerToStringTabAddTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
                txtVisitorToStringTabAddTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
            Else 'FOUND

                Dim age As Integer = foundValue.calcAge(foundValue.Birthdate, foundValue.DatePurchased)

                txtCustomerToStringTabAddTabPassportFeaturesTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.Owner.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf
                txtVisitorToStringTabAddTabPassportFeaturesTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & vbCrLf _
                        & "Name: " & foundValue.VisitorName & vbCrLf _
                        & "Date Purchased: " & foundValue.DatePurchased & vbCrLf _
                        & "Birthdate: " & Format(foundValue.Birthdate, "yyyy") & Format(foundValue.Birthdate, "MM") & Format(foundValue.Birthdate, "dd") & vbCrLf _
                        & "Age: " & age.ToString & vbCrLf _
                        & "Is Child - " & foundValue.isChildUnder13(age) & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf

                cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = featureSelectedIndex

                '& "Age: " & foundValue.Age & vbCrLf _
                '& "Is Child: " & foundValue.IsChild & vbCrLf _


                'End If

            End If

        End If

    End Sub '_cboPassportIdTabAddTabPassportFeaturesTbcMain_SelectedIndexChanged



    Private Sub _cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedIndexChanged

        txtCustomerToStringTabUpdateTabPassportFeaturesTbcMain.Text = ""
        txtVisitorToStringTabUpdateTabPassportFeaturesTbcMain.Text = ""

        If cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedIndex

            Dim inputValue As String
            Dim foundValue As PassportFeature
            Dim foundLocation As Integer

            'get/validate input

            inputValue = cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.SelectedItem.ToString

            'do processing

            foundValue = _theThemePark.findPassportFeature(inputValue, foundLocation)



            'display info

            If foundValue Is Nothing Then 'NOT found
                txtCustomerToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
                txtVisitorToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
                txtFeatureToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValue & "' NOT found."
                txtPriceGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text &=
                     "N/A"
                txtCurrentQuantityGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text &=
                     "N/A"

            Else 'FOUND

                Dim age As Integer = foundValue.Passport.calcAge(foundValue.Passport.Birthdate, foundValue.Passport.DatePurchased)

                txtCustomerToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & foundValue.Passport.Owner.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf
                txtVisitorToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                         "'" & inputValue & "' FOUND: " & vbCrLf _
                        & "Name: " & foundValue.Passport.VisitorName & vbCrLf _
                        & "Date Purchased: " & Format(foundValue.Passport.DatePurchased, "yyyy") & Format(foundValue.Passport.DatePurchased, "MM") & Format(foundValue.Passport.DatePurchased, "dd") & vbCrLf _
                        & "Birthdate: " & Format(foundValue.Passport.Birthdate, "yyyy") & Format(foundValue.Passport.Birthdate, "MM") & Format(foundValue.Passport.Birthdate, "dd") & vbCrLf _
                        & "Age: " & age.ToString & vbCrLf _
                        & "Is Child - " & foundValue.Passport.isChildUnder13(age) & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf
                txtFeatureToStringTabUpdateTabPassportFeaturesTbcMain.Text &=
                      "'" & inputValue & "' FOUND: " & foundValue.Feature.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocation.ToString _
                        & "." & vbCrLf
                txtPriceGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text =
                    foundValue.Price.ToString
                txtCurrentQuantityGrpUpdateTabUpdateTabPassportFeaturesTbcMain.Text =
                    foundValue.Quantity.ToString

            End If
        End If


    End Sub '_cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain_SelectedIndexChanged

    Private Sub _cboFeatureIdTabAddTabPassportFeaturesTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndexChanged,
        cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndexChanged

        txtFeatureToStringTabAddTabPassportFeaturesTbcMain.Text = ""



        If cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex >= 0 Then

            Dim selectedIndex As Integer = cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedIndex
            Dim inputValueFeature As String
            Dim foundValueFeature As Feature
            Dim foundLocationFeature As Integer

            'get/validate input

            inputValueFeature = cboFeatureIdTabAddTabPassportFeaturesTbcMain.SelectedItem.ToString

            'do processing

            foundValueFeature = _theThemePark.findFeature(inputValueFeature, foundLocationFeature)

            Dim inputValuePassport As String
            Dim foundValuePassport As Passport
            Dim foundLocationPassport As Integer

            'get/validate input
            If cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex < 0 Then
                cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedIndex = 0
            End If

            inputValuePassport = cboPassportIdTabAddTabPassportFeaturesTbcMain.SelectedItem.ToString

            'do processing

            foundValuePassport = _theThemePark.findPassport(inputValuePassport, foundLocationPassport)


            'display info

            If foundValueFeature Is Nothing Then 'NOT found
                txtFeatureToStringTabAddTabPassportFeaturesTbcMain.Text &=
                     "'" & inputValueFeature & "' NOT found."
                txtAdultPriceGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text = "NO PRICE"

            Else 'FOUND
                txtFeatureToStringTabAddTabPassportFeaturesTbcMain.Text &=
                         "'" & inputValueFeature & "' FOUND: " & foundValueFeature.ToString & vbCrLf & vbCrLf _
                        & "Found at index: " & foundLocationFeature.ToString _
                        & "." & vbCrLf

                If foundValuePassport Is Nothing Then
                    txtAdultPriceGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text = "N/A"

                Else

                    Dim age As Integer = foundValuePassport.calcAge(foundValuePassport.Birthdate, foundValuePassport.DatePurchased)

                    If foundValuePassport.isChildUnder13(age) Then
                        txtAdultPriceGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text = foundValueFeature.ChildPricePerUnit.ToString
                    Else
                        txtAdultPriceGrpPriceQuantityTabAddTabPassportFeaturesTbcMain.Text = foundValueFeature.AdultPricePerUnit.ToString
                    End If

                End If

            End If

        End If


    End Sub '_cboFeatureIdTabAddTabPassportFeaturesTbcMain_SelectedIndexChanged



    Private Sub _cboPassportIdTabUseFeatureTbcMain_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) _
        Handles cboPassportIdTabUseFeatureTbcMain.SelectedIndexChanged

        If cboPassportIdTabUseFeatureTbcMain.SelectedIndex >= 0 Then
            Dim selectedIndex As Integer = cboPassportIdTabUseFeatureTbcMain.SelectedIndex

            txtCustomerToStringTabUseFeatureTbcMain.Text = ""
            txtVisitorToStringTabUseFeatureTbcMain.Text = ""
            txtFeatureToStringTabUseFeatureTbcMain.Text = ""
            txtPreviouslyUsedTabUseFeatureTbcMain.Text = ""
            txtRemainingQuantityGrpFeatureDetailsTabUseFeatureTbcMain.Text = ""

            If cboPassportIdTabUseFeatureTbcMain.SelectedIndex >= 0 Then


                Dim inputValue As String
                Dim foundValue As PassportFeature
                Dim foundLocation As Integer

                'get/validate input

                inputValue = cboPassportIdTabUseFeatureTbcMain.SelectedItem.ToString

                'do processing

                foundValue = _theThemePark.findPassportFeature(inputValue, foundLocation)



                'display info

                If foundValue Is Nothing Then 'NOT found
                    txtCustomerToStringTabUseFeatureTbcMain.Text &=
                         "'" & inputValue & "' NOT found."
                    txtVisitorToStringTabUseFeatureTbcMain.Text &=
                         "'" & inputValue & "' NOT found."
                    txtFeatureToStringTabUseFeatureTbcMain.Text &=
                         "'" & inputValue & "' NOT found."
                    txtRemainingQuantityGrpFeatureDetailsTabUseFeatureTbcMain.Text &=
                         "N/A"


                Else 'FOUND

                    Dim age As Integer = foundValue.Passport.calcAge(foundValue.Passport.Birthdate, foundValue.Passport.DatePurchased)

                    txtCustomerToStringTabUseFeatureTbcMain.Text &=
                             "'" & inputValue & "' FOUND: " & foundValue.Passport.Owner.ToString & vbCrLf & vbCrLf _
                            & "Found at index: " & foundLocation.ToString _
                            & "." & vbCrLf
                    txtVisitorToStringTabUseFeatureTbcMain.Text &=
                             "'" & inputValue & "' FOUND: " & vbCrLf _
                            & "Name: " & foundValue.Passport.VisitorName & vbCrLf _
                            & "Date Purchased: " & Format(foundValue.Passport.DatePurchased, "yyyy") & Format(foundValue.Passport.DatePurchased, "MM") & Format(foundValue.Passport.DatePurchased, "dd") & vbCrLf _
                            & "Birthdate: " & Format(foundValue.Passport.Birthdate, "yyyy") & Format(foundValue.Passport.Birthdate, "MM") & Format(foundValue.Passport.Birthdate, "dd") & vbCrLf _
                            & "Age: " & age.ToString & vbCrLf _
                            & "Is Child - " & foundValue.Passport.isChildUnder13(age) & vbCrLf & vbCrLf _
                            & "Found at index: " & foundLocation.ToString _
                            & "." & vbCrLf
                    txtFeatureToStringTabUseFeatureTbcMain.Text &=
                          "'" & inputValue & "' FOUND: " & foundValue.Feature.ToString & vbCrLf & vbCrLf _
                            & "Found at index: " & foundLocation.ToString _
                            & "." & vbCrLf
                    txtRemainingQuantityGrpFeatureDetailsTabUseFeatureTbcMain.Text =
                        foundValue.Quantity.ToString

                End If
            End If

            For Each item As UsedFeature In _theThemePark.iterateUsedFeature()

                If item Is Nothing Then
                    txtPreviouslyUsedTabUseFeatureTbcMain.Text = " nothing"
                Else
                    If item.PassportFeature.id = cboPassportIdTabUseFeatureTbcMain.SelectedItem.ToString Then

                        txtPreviouslyUsedTabUseFeatureTbcMain.Text &= vbCrLf _
                                & Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") & " " _
                                & ", Location: " & item.LocationWhereUsed.ToString & " " _
                                & ", Quantity Used: " & item.QuantityUsed.ToString


                    End If
                End If




            Next item

        End If

    End Sub '_cboPassportIdTabUseFeatureTbcMain_SelectedIndexChanged



    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

    Private Sub _customerAdded(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_CustomerAdded

        Dim theThemePark_EventArgs_CustomerAdded As ThemePark_EventArgs_CustomerAdded
        Dim theCustomer As Customer

        theThemePark_EventArgs_CustomerAdded = CType(e, ThemePark_EventArgs_CustomerAdded)

        theCustomer = theThemePark_EventArgs_CustomerAdded.Customer

        With theCustomer
            lstCustomerListTabSummaryInfoTbcMain.Items.Add(.id.ToString)
            cboCustomerNameTabPassportsTbcMain.Items.Add(.id.ToString)
        End With

        txtTransactions.Text &= vbCrLf & "Customer Created And Initalized! " & theCustomer.ToString & vbCrLf
        txtQtyCustomerListTabSummaryInfoTbcMain.Text = _theThemePark.NumberCustomers.ToString


        _refreshMetrics()



    End Sub '_customerAdded

    Private Sub _PassportAdded(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_PassportAdded

        Dim theThemePark_EventArgs_PassportAdded As ThemePark_EventArgs_PassportAdded
        Dim thePassport As Passport

        theThemePark_EventArgs_PassportAdded = CType(e, ThemePark_EventArgs_PassportAdded)

        thePassport = theThemePark_EventArgs_PassportAdded.Passport

        With thePassport
            lstPassportListTabSummaryInfoTbcMain.Items.Add(.id.ToString)
            cboPassportIdTabAddTabPassportFeaturesTbcMain.Items.Add(.id.ToString)

        End With

        txtTransactions.Text &= vbCrLf & "Passport Created And Initalized! " & thePassport.ToString & vbCrLf


        txtQtyPassPortListTabSummaryInfoTbcMain.Text = _theThemePark.NumberPassports.ToString
        _refreshMetrics()


    End Sub '_PassportAdded

    Private Sub _FeatureAdded(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_FeatureAdded

        Dim theThemePark_EventArgs_FeatureAdded As ThemePark_EventArgs_FeatureAdded
        Dim theFeature As Feature

        theThemePark_EventArgs_FeatureAdded = CType(e, ThemePark_EventArgs_FeatureAdded)

        theFeature = theThemePark_EventArgs_FeatureAdded.Feature

        With theFeature
            lstFeatureListTabSummaryInfoTbcMain.Items.Add(.id.ToString)
            cboFeatureIdTabAddTabPassportFeaturesTbcMain.Items.Add(.id.ToString)
        End With

        txtTransactions.Text &= vbCrLf & "Feature Created And Initalized! " & theFeature.ToString & vbCrLf


        txtQtyFeatureListTabSummaryInfoTbcMain.Text = _theThemePark.NumberFeatures.ToString
        _refreshMetrics()

    End Sub '_FeatureAdded

    Private Sub _PassportFeatureAdded(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_PassportFeatureAdded

        Dim theThemePark_EventArgs_PassportFeatureAdded As ThemePark_EventArgs_PassportFeatureAdded
        Dim thePassportFeature As PassportFeature

        theThemePark_EventArgs_PassportFeatureAdded = CType(e, ThemePark_EventArgs_PassportFeatureAdded)

        thePassportFeature = theThemePark_EventArgs_PassportFeatureAdded.PassportFeature

        With thePassportFeature
            lstPassportFeatureListTabSummaryInfoTbcMain.Items.Add(.id.ToString)
            cboPassportFeatureIdTabUpdateTabPassportFeaturesTbcMain.Items.Add(.id.ToString)
            cboPassportIdTabUseFeatureTbcMain.Items.Add(.id.ToString)
        End With


        txtTransactions.Text &= "Passport Feature Created and Initalized! " & thePassportFeature.ToString & vbCrLf



        txtQtyPassportFeatureListTabSummaryInfoTbcMain.Text = _theThemePark.NumberPassportFeatures.ToString
        _refreshMetrics()


    End Sub '_PassportFeatureAdded

    Private Sub _UpdatePassportFeature(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_UpdatePassportFeature

        Dim theThemePark_EventArgs_UpdatePassportFeature As ThemePark_EventArgs_UpdatePassportFeature
        Dim thePassportFeature As PassportFeature

        theThemePark_EventArgs_UpdatePassportFeature = CType(e, ThemePark_EventArgs_UpdatePassportFeature)

        thePassportFeature = theThemePark_EventArgs_UpdatePassportFeature.PassportFeature

        With thePassportFeature

        End With

        txtTransactions.Text &= vbCrLf & "Passport Feature updated! " & thePassportFeature.ToString & vbCrLf

        _refreshMetrics()


    End Sub '_UpdatePassportFeature

    Private Sub _UsedFeature(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mThemePark.ThemePark_UsedFeature

        Dim theThemePark_EventArgs_UsedFeature As ThemePark_EventArgs_UsedFeature
        Dim theUsedFeature As UsedFeature

        theThemePark_EventArgs_UsedFeature = CType(e, ThemePark_EventArgs_UsedFeature)

        theUsedFeature = theThemePark_EventArgs_UsedFeature.UsedFeature

        With theUsedFeature
            lstUsedFeatureListTabSummaryInfoTbcMain.Items.Add(.id.ToString)
            ' theUsedFeature.PassportFeature.QuantityRemaining -= theUsedFeature.QuantityUsed
        End With

        txtTransactions.Text &= vbCrLf & "Used Feature Created And Initalized! " & theUsedFeature.ToString & vbCrLf


        txtQtyUsedFeatureListTabSummaryInfoTbcMain.Text = _theThemePark.NumberUsedFeatures.ToString
        _refreshMetrics()


    End Sub '_UsedFeature




#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'These are all public.



#End Region 'Events

End Class 'FrmMain