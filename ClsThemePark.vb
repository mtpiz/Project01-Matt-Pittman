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


'Tier:               Business Logic
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

Public Class ThemePark

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    Private Const mMAX As Integer = 5
    Private Const mARRAY_SIZE_DEFAULT As Integer = 5
    Private Const mARRAY_INCREMENT_DEFAULT As Integer = 5

    '********** Module-level variables

    Private mThemeParkName As String
    Private mNumberCustomers As Integer
    Private mNumberPassports As Integer
    Private mNumberFeatures As Integer
    Private mNumberPassportFeatures As Integer
    Private mNumberUsedFeatures As Integer

    Private mCustomer() As Customer
    Private mMaxCustomers As Integer
    Private mNumCustomers As Integer

    Private mFeature() As Feature
    Private mMaxFeatures As Integer
    Private mNumFeatures As Integer

    Private mPassport() As Passport
    Private mMaxPassports As Integer
    Private mNumPassports As Integer

    Private mPassportFeature() As PassportFeature
    Private mMaxPassportFeatures As Integer
    Private mNumPassportFeatures As Integer

    Private mUsedFeature() As UsedFeature
    Private mMaxUsedFeatures As Integer
    Private mNumUsedFeatures As Integer

    Private mTransactions() As String
    Private mMaxTransactions As Integer
    Private mNumTransactions As Integer

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'These are all public.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes
    Public Sub New(
            ByVal pThemeParkName As String
            )

        MyBase.New()
        _ThemeParkName = pThemeParkName


        ReDim mCustomer(_MAX - 1)
        _MaxCustomers = _ARRAY_SIZE_DEFAULT
        ReDim mCustomer(_MaxCustomers - 1)
        _numCustomers = 0

        ReDim mFeature(_MAX - 1)
        _MaxFeatures = _ARRAY_SIZE_DEFAULT
        ReDim mFeature(_MaxFeatures - 1)
        _numFeatures = 0

        ReDim mPassport(_MAX - 1)
        _MaxPassports = _ARRAY_SIZE_DEFAULT
        ReDim mPassport(_MaxPassports - 1)
        _numPassports = 0

        ReDim mPassportFeature(_MAX - 1)
        _MaxPassportFeatures = _ARRAY_SIZE_DEFAULT
        ReDim mPassportFeature(_MaxPassportFeatures - 1)
        _numCustomers = 0

        ReDim mUsedFeature(_MAX - 1)
        _MaxUsedFeatures = _ARRAY_SIZE_DEFAULT
        ReDim mUsedFeature(_MaxUsedFeatures - 1)
        _numUsedFeatures = 0

        ReDim mTransactions(_MAX - 1)
        _MaxTransactions = _ARRAY_SIZE_DEFAULT
        ReDim mTransactions(_MaxTransactions - 1)
        _numTransactions = 0


    End Sub 'New(pId,pName)


    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public Property ThemeParkName As String
        Get
            Return _ThemeParkName()
        End Get
        Set(pValue As String)
            _ThemeParkName() = pValue
        End Set
    End Property

    Public ReadOnly Property NumberCustomers As Integer
        Get
            Return _NumberCustomers()
        End Get
        'Set(pValue As Integer)
        ' _NumberCustomers() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumberPassports As Integer
        Get
            Return _NumberPassports()
        End Get
        'Set(pValue As Integer)
        '_NumberPassports() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumberFeatures As Integer
        Get
            Return _NumberFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumberPassportFeatures As Integer
        Get
            Return _NumberPassportFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberPassportFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumberUsedFeatures As Integer
        Get
            Return _NumberUsedFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxCustomers As Integer
        Get
            Return _MaxCustomers()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumCustomers As Integer
        Get
            Return _numCustomers()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxFeatures As Integer
        Get
            Return _MaxFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumFeatures As Integer
        Get
            Return _numFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxPassports As Integer
        Get
            Return _MaxPassports()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumPassports As Integer
        Get
            Return _numPassports()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxPassportFeatures As Integer
        Get
            Return _MaxPassportFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumPassportFeatures As Integer
        Get
            Return _numPassportFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxUsedFeatures As Integer
        Get
            Return _MaxUsedFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumUsedFeatures As Integer
        Get
            Return _numUsedFeatures()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property MaxTransactions As Integer
        Get
            Return _MaxTransactions()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public ReadOnly Property NumTransactions As Integer
        Get
            Return _numTransactions()
        End Get
        'Set(pValue As Integer)
        '_NumberUsedFeatures() = pValue
        'End Set
    End Property

    Public Iterator Function iterateCustomer(
        ) _
    As _
        IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateCustomer()
            Yield theObject
        Next theObject

    End Function '_iterateItem()
    Public Iterator Function iteratePassport(
        ) _
    As _
        IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iteratePassport()
            Yield theObject
        Next theObject

    End Function '_iterateItem()
    Public Iterator Function iterateFeature(
        ) _
    As _
        IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateFeature()
            Yield theObject
        Next theObject

    End Function '_iterateItem()
    Public Iterator Function iteratePassportFeature(
        ) _
    As _
        IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iteratePassportFeature()
            Yield theObject
        Next theObject

    End Function '_iterateItem()
    Public Iterator Function iterateUsedFeature(
        ) _
    As _
        IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateUsedFeature()
            Yield theObject
        Next theObject

    End Function '_iterateItem()

    Public Iterator Function iterateTransactions(
        ) _
    As _
        IEnumerable(Of String)

        Dim theObject As String

        For Each theObject In _iterateTransactions()
            Yield theObject
        Next theObject

    End Function '_iterateItem()



    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _ThemeParkName() As String
        Get
            Return mThemeParkName
        End Get
        Set(ByVal pValue As String)
            mThemeParkName = pValue
        End Set
    End Property

    Private Property _NumberCustomers() As Integer
        Get
            Return mNumberCustomers
        End Get
        Set(ByVal pValue As Integer)
            mNumberCustomers = pValue
        End Set
    End Property

    Private Property _NumberPassports() As Integer
        Get
            Return mNumberPassports
        End Get
        Set(ByVal pValue As Integer)
            mNumberPassports = pValue
        End Set
    End Property

    Private Property _NumberFeatures() As Integer
        Get
            Return mNumberFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumberFeatures = pValue
        End Set
    End Property

    Private Property _NumberPassportFeatures() As Integer
        Get
            Return mNumberPassportFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumberPassportFeatures = pValue
        End Set
    End Property

    Private Property _NumberUsedFeatures() As Integer
        Get
            Return mNumberUsedFeatures
        End Get
        Set(ByVal pValue As Integer)
            mNumberUsedFeatures = pValue
        End Set
    End Property

    Private ReadOnly Property _MAX As Integer
        Get
            Return mMAX
        End Get
    End Property

    Private ReadOnly Property _ARRAY_SIZE_DEFAULT As Integer
        Get
            Return mARRAY_SIZE_DEFAULT
        End Get
    End Property

    Private ReadOnly Property _ARRAY_INCREMENT_DEFAULT As Integer
        Get
            Return mARRAY_INCREMENT_DEFAULT
        End Get
    End Property

    Private Property _MaxCustomers As Integer
        Get
            Return mMaxCustomers
        End Get
        Set(pValue As Integer)
            mMaxCustomers = pValue
        End Set
    End Property

    Private Property _numCustomers As Integer
        Get
            Return mNumCustomers
        End Get
        Set(pValue As Integer)
            mNumCustomers = pValue
        End Set
    End Property
    Private Property _MaxFeatures As Integer
        Get
            Return mMaxFeatures
        End Get
        Set(pValue As Integer)
            mMaxFeatures = pValue
        End Set
    End Property

    Private Property _numFeatures As Integer
        Get
            Return mNumFeatures
        End Get
        Set(pValue As Integer)
            mNumFeatures = pValue
        End Set
    End Property
    Private Property _MaxPassports As Integer
        Get
            Return mMaxPassports
        End Get
        Set(pValue As Integer)
            mMaxPassports = pValue
        End Set
    End Property

    Private Property _numPassports As Integer
        Get
            Return mNumPassports
        End Get
        Set(pValue As Integer)
            mNumPassports = pValue
        End Set
    End Property
    Private Property _MaxPassportFeatures As Integer
        Get
            Return mMaxPassportFeatures
        End Get
        Set(pValue As Integer)
            mMaxPassportFeatures = pValue
        End Set
    End Property

    Private Property _numPassportFeatures As Integer
        Get
            Return mNumPassportFeatures
        End Get
        Set(pValue As Integer)
            mNumPassportFeatures = pValue
        End Set
    End Property
    Private Property _MaxUsedFeatures As Integer
        Get
            Return mMaxUsedFeatures
        End Get
        Set(pValue As Integer)
            mMaxUsedFeatures = pValue
        End Set
    End Property

    Private Property _numUsedFeatures As Integer
        Get
            Return mNumUsedFeatures
        End Get
        Set(pValue As Integer)
            mNumUsedFeatures = pValue
        End Set
    End Property

    Private Property _MaxTransactions As Integer
        Get
            Return mMaxTransactions
        End Get
        Set(pValue As Integer)
            mMaxTransactions = pValue
        End Set
    End Property

    Private Property _numTransactions As Integer
        Get
            Return mNumTransactions
        End Get
        Set(pValue As Integer)
            mNumTransactions = pValue
        End Set
    End Property

    Private Iterator Function _iterateCustomer(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numCustomers - 1
            Yield _ithCustomer(i)
        Next i

    End Function '_iterateItem()
    Private Iterator Function _iteratePassport(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numPassports - 1
            Yield _ithPassport(i)
        Next i

    End Function '_iterateItem()
    Private Iterator Function _iterateFeature(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numFeatures - 1
            Yield _ithFeature(i)
        Next i

    End Function '_iterateItem()
    Private Iterator Function _iteratePassportFeature(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numPassportFeatures - 1
            Yield _ithPassportFeature(i)
        Next i

    End Function '_iterateItem()
    Private Iterator Function _iterateUsedFeature(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numUsedFeatures - 1
            Yield _ithUsedFeature(i)
        Next i

    End Function '_iterateItem()

    Private Iterator Function _iterateTransactions(
            ) _
        As _
            IEnumerable(Of String)

        Dim i As Integer

        For i = 0 To _numTransactions - 1
            Yield _ithTransactions(i)
        Next i

    End Function '_iterateItem()



    Private Property _ithCustomer(ByVal pN As Integer) As Customer
        'Assumes: 0 <= pN < _numCustomers.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxCustomers Then
                Return mCustomer(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As Customer)
            If pN >= 0 And pN < _MaxCustomers Then
                mCustomer(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithFeature(ByVal pN As Integer) As Feature
        'Assumes: 0 <= pN < _numFeatures.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxFeatures Then
                Return mFeature(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As Feature)
            If pN >= 0 And pN < _MaxFeatures Then
                mFeature(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithPassport(ByVal pN As Integer) As Passport
        'Assumes: 0 <= pN < _numPassports.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxPassports Then
                Return mPassport(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As Passport)
            If pN >= 0 And pN < _MaxPassports Then
                mPassport(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithPassportFeature(ByVal pN As Integer) As PassportFeature
        'Assumes: 0 <= pN < _numPassportFeatures.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxPassportFeatures Then
                Return mPassportFeature(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As PassportFeature)
            If pN >= 0 And pN < _MaxPassportFeatures Then
                mPassportFeature(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithUsedFeature(ByVal pN As Integer) As UsedFeature
        'Assumes: 0 <= pN < _numUsedFeatures.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxUsedFeatures Then
                Return mUsedFeature(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As UsedFeature)
            If pN >= 0 And pN < _MaxUsedFeatures Then
                mUsedFeature(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithTransactions(ByVal pN As Integer) As String
        'Assumes: 0 <= pN < _numUsedFeatures.
        'Throws an IndexOutOfRangeException if this is not the case.
        Get
            If pN >= 0 And pN < _MaxTransactions Then
                Return mTransactions(pN)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As String)
            If pN >= 0 And pN < _MaxTransactions Then
                mTransactions(pN) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property


#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Function addCustomer(
          ByVal pID As String,
          ByVal pName As String
          ) _
      As _
          Customer

        Return _addCustomer(pID, pName)

    End Function 'addCustomer

    Public Function addFeature(
            ByVal pId As String,
            ByVal pName As String,
            ByVal pUnitOfMeasure As String,
            ByVal pAdultPricePerUnit As Decimal,
            ByVal pChildPricePerUnit As Decimal
          ) _
      As _
          Feature

        Return _addFeature(pId, pName, pUnitOfMeasure, pAdultPricePerUnit, pChildPricePerUnit)


    End Function 'addFeature

    Public Function addPassport(
            ByVal pId As String,
            ByVal pDatePurchased As Date,
            ByVal pVisitorName As String,
            ByVal pBirthdate As Date,
            ByVal pOwner As Customer
          ) _
      As _
          Passport

        Return _addPassport(pId, pDatePurchased, pVisitorName, pBirthdate, pOwner)


    End Function 'addFeature




    Public Function addPassportFeature(
            ByVal pId As String,
            ByVal pQuantityPurchased As Decimal,
            ByVal pPrice As Decimal,
            ByVal pPassport As Passport,
            ByVal pFeature As Feature
            ) _
      As _
          PassportFeature

        Return _addPassportFeature(pId, pQuantityPurchased, pPrice, pPassport, pFeature)

    End Function 'addPassportFeature

    Public Function updatePassportFeature(
            ByVal pId As String,
            ByVal pQuantityPurchased As Decimal,
            ByVal pPrice As Decimal,
            ByVal pPassportFeature As PassportFeature,
            ByVal pPassport As Passport,
            ByVal pFeature As Feature,
            ByVal pDateUpdated As Date
          ) _
      As _
          PassportFeature

        Return _updatePassportFeature(pId, pQuantityPurchased, pPrice, pPassportFeature, pPassport, pFeature, pDateUpdated)

    End Function

    Public Function addUsedFeature(
            ByVal pId As String,
            ByVal pDateUsed As Date,
            ByVal pLocationWhereUsed As String,
            ByVal pQuantityUsed As Decimal,
            ByVal pPassportFeature As PassportFeature
            ) _
      As _
          UsedFeature

        Return _addUsedFeature(pId, pDateUsed, pLocationWhereUsed, pQuantityUsed, pPassportFeature)

    End Function 'addPassportFeature

    Public Function findCustomer(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Customer

        Return _findCustomer(pItemToFind, pLocationFound)

    End Function 'findItem(pItemToFind)

    Public Function findFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Feature

        Return _findFeature(pItemToFind, pLocationFound)

    End Function 'findItem(pItemToFind)

    Public Function findPassportFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            PassportFeature

        Return _findPassportFeature(pItemToFind, pLocationFound)

    End Function 'findItem(pItemToFind)

    Public Function findPassport(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Passport

        Return _findPassport(pItemToFind, pLocationFound)

    End Function 'findItem(pItemToFind)

    Public Function findUsedFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            UsedFeature

        Return _findUsedFeature(pItemToFind, pLocationFound)

    End Function 'findItem(pItemToFind)
    Public Sub readFile()

        _readFile()

    End Sub

    Public Sub writeFile(ByRef pAppendFile As Boolean)

        _writeFile(pAppendFile)

    End Sub

    Public Function averageBalanceUnused() As String

        Return _averageBalanceUnused()

    End Function

    Public Function sumUnusedPassportFeature() As String

        Return _sumUnusedPassportFeature()

    End Function
    Public Function averagePassportsPerCustomer() As String

        Return _averagePassportsPerCustomer()

    End Function


    Public Function mostPopularPassportFeature() As String

        Return _mostPopularPassportFeature()

    End Function

    Public Function percentPassportFeaturesUsed() As String

        Return _percentPassportFeaturesUsed()

    End Function

    Public Function averageAge() As String

        Return _averageAge()

    End Function
    Public Function passportHoldersBirthday() As String

        Return _passportHoldersBirthday()

    End Function



    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()


    '********** Private Non-Shared Behavioral Methods

    Private Function _addCustomer(
          ByVal pID As String,
          ByVal pName As String
          ) _
      As _
          Customer

        'declare variables
        Dim theCustomer As Customer
        theCustomer = New Customer(pID, pName)
        Dim theCurrentTime As Date
        Dim customerTransaction As String

        'get/validate input
        If _numCustomers >= _MaxCustomers Then
            _MaxCustomers += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mCustomer(_MaxCustomers - 1)
        End If

        'do processing
        Try
            _ithCustomer(_numCustomers) = theCustomer
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numCustomers += 1
        _NumberCustomers += 1

        'display output

        'Add transaction to trans array
        theCurrentTime = Now
        customerTransaction =
            Format(Now, "yyyy").ToString &
            Format(Now, "MM").ToString &
            Format(Now, "dd").ToString &
            "; " & Format(Now, "hh").ToString &
            Format(Now, "mm").ToString &
            "; CUSTOMER; CREATE; " &
            theCustomer.id.ToString &
            "; " & theCustomer.CustomerName.ToString

        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = customerTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1


        RaiseEvent ThemePark_CustomerAdded(Me, New ThemePark_EventArgs_CustomerAdded(theCustomer)) 'RaiseEvent

        Return theCustomer

        'get ready for next input

    End Function '_addCustomer


    Private Function _addFeature(
            ByVal pId As String,
            ByVal pName As String,
            ByVal pUnitOfMeasure As String,
            ByVal pAdultPricePerUnit As Decimal,
            ByVal pChildPricePerUnit As Decimal
          ) _
      As _
          Feature

        Dim theFeature As Feature
        Dim theCurrentTime As DateTime
        Dim featureTransaction As String
        theFeature = New Feature(pId, pName, pUnitOfMeasure, pAdultPricePerUnit, pChildPricePerUnit)

        If _numFeatures >= _MaxFeatures Then
            _MaxFeatures += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mFeature(_MaxFeatures - 1)
        End If

        'do processing
        Try
            _ithFeature(_numFeatures) = theFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numFeatures += 1
        _NumberFeatures += 1

        'do processing


        'Add transaction to trans array
        theCurrentTime = Now
        featureTransaction =
            Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") &
            "; FEATURE; CREATE; " &
            theFeature.id.ToString &
            "; " & theFeature.Name &
            "; " & theFeature.UnitOfMeasure &
            "; " & theFeature.AdultPricePerUnit &
            "; " & theFeature.ChildPricePerUnit

        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = featureTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1

        RaiseEvent ThemePark_FeatureAdded(Me, New ThemePark_EventArgs_FeatureAdded(theFeature)) 'RaiseEvent

        Return theFeature


    End Function '_addFeature

    Private Function _addPassport(
            ByVal pId As String,
            ByVal pDatePurchased As Date,
            ByVal pVisitorName As String,
            ByVal pBirthdate As Date,
            ByVal pOwner As Customer
          ) _
      As _
          Passport

        Dim thePassport As Passport
        Dim theCurrentTime As DateTime
        Dim PassportTransaction As String

        thePassport = New Passport(pId, pDatePurchased, pVisitorName, pBirthdate, pOwner)


        If _numPassports >= _MaxPassports Then
            _MaxPassports += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mPassport(_MaxPassports - 1)
        End If

        'do processing
        Try
            _ithPassport(_numPassports) = thePassport
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numPassports += 1
        _NumberPassports += 1

        'Add transaction to trans array
        theCurrentTime = Now
        PassportTransaction =
            Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") &
            "; PASSBOOK; CREATE; " &
            thePassport.id.ToString &
            "; " & thePassport.Owner.id.ToString &
            "; " & Format(thePassport.DatePurchased, "yyyy").ToString & Format(thePassport.DatePurchased, "MM").ToString & Format(thePassport.DatePurchased, "dd").ToString &
            "; " & thePassport.VisitorName.ToString &
            "; " & Format(thePassport.Birthdate, "yyyy").ToString & Format(thePassport.Birthdate, "MM").ToString & Format(thePassport.Birthdate, "dd").ToString

        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = PassportTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1

        RaiseEvent ThemePark_PassportAdded(Me, New ThemePark_EventArgs_PassportAdded(thePassport)) 'RaiseEvent

        Return thePassport


    End Function 'addFeature

    Private Function _addPassportFeature(
            ByVal pId As String,
            ByVal pQuantityPurchased As Decimal,
            ByVal pPrice As Decimal,
            ByVal pPassport As Passport,
            ByVal pFeature As Feature
          ) _
      As _
          PassportFeature

        Dim thePassportFeature As PassportFeature
        Dim theCurrentTime As DateTime
        Dim passportFeatureTransaction As String
        thePassportFeature = New PassportFeature(pId, pQuantityPurchased, pPrice, pPassport, pFeature)

        If _numPassportFeatures >= _MaxPassportFeatures Then
            _MaxPassportFeatures += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mPassportFeature(_MaxPassportFeatures - 1)
        End If

        'do processing
        Try
            _ithPassportFeature(_numPassportFeatures) = thePassportFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numPassportFeatures += 1
        _NumberPassportFeatures += 1

        ''Add transaction to trans array
        theCurrentTime = Now
        passportFeatureTransaction =
            Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") &
            "; PASSBOOK_FEATURE; PURCHASE; " &
            thePassportFeature.id.ToString &
            "; " & thePassportFeature.Quantity &
            "; " & thePassportFeature.Passport.id &
            "; " & thePassportFeature.Feature.id


        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = passportFeatureTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1


        RaiseEvent ThemePark_PassportFeatureAdded(Me, New ThemePark_EventArgs_PassportFeatureAdded(thePassportFeature)) 'RaiseEvent

        Return thePassportFeature

    End Function '_addPassportFeature

    Private Function _updatePassportFeature(
            ByVal pId As String,
            ByVal pQuantity As Decimal,
            ByVal pPrice As Decimal,
            ByVal pPassportFeature As PassportFeature,
            ByVal pPassport As Passport,
            ByVal pFeature As Feature,
            ByVal pDateUpdated As Date
          ) _
      As _
          PassportFeature

        Dim theCurrentTime As DateTime
        Dim updatePassportTransaction As String

        ''Add transaction to trans array
        theCurrentTime = Now
        updatePassportTransaction =
            Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") &
            "; PASSBOOK_FEATURE; UPDATE; " &
            pPassportFeature.id.ToString &
            "; " & Format(pDateUpdated, "yyyy") & Format(pDateUpdated, "MM") & Format(pDateUpdated, "dd") &
            "; " & pQuantity.ToString



        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = updatePassportTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1

        RaiseEvent ThemePark_UpdatePassportFeature(Me, New ThemePark_EventArgs_UpdatePassportFeature(pPassportFeature)) 'RaiseEvent

        pPassportFeature.Quantity = pQuantity

        Return pPassportFeature

    End Function '_addPassportFeature



    Private Function _addUsedFeature(
            ByVal pId As String,
            ByVal pDateUsed As Date,
            ByVal pLocationWhereUsed As String,
            ByVal pQuantityUsed As Decimal,
            ByVal pPassportFeature As PassportFeature
            ) _
      As _
          UsedFeature

        Dim theUsedFeature As UsedFeature
        Dim theCurrentTime As DateTime
        Dim usedFeatureTransaction As String
        theUsedFeature = New UsedFeature(pId, pDateUsed, pLocationWhereUsed, pQuantityUsed, pPassportFeature)

        If _numUsedFeatures >= _MaxUsedFeatures Then
            _MaxUsedFeatures += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mUsedFeature(_MaxUsedFeatures - 1)
        End If

        'do processing
        Try
            _ithUsedFeature(_numUsedFeatures) = theUsedFeature
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numUsedFeatures += 1
        _NumberUsedFeatures += 1

        ''Add transaction to trans array
        theCurrentTime = Now
        usedFeatureTransaction =
            Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") & "; " & Format(Now, "hh") & Format(Now, "mm") &
            "; PASSBOOK_FEATURE; USE; " &
            theUsedFeature.id.ToString &
            "; " & theUsedFeature.PassportFeature.id &
            "; " & Format(Now, "yyyy") & Format(Now, "MM") & Format(Now, "dd") &
            "; " & theUsedFeature.LocationWhereUsed &
            "; " & theUsedFeature.QuantityUsed



        If _numTransactions >= _MaxTransactions Then
            _MaxTransactions += _ARRAY_INCREMENT_DEFAULT
            ReDim Preserve mTransactions(_MaxTransactions - 1)
        End If

        'do processing
        Try
            _ithTransactions(_numTransactions) = usedFeatureTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try
        _numTransactions += 1

        pPassportFeature.Quantity = pPassportFeature.Quantity - pQuantityUsed

        RaiseEvent ThemePark_UsedFeature(Me, New ThemePark_EventArgs_UsedFeature(theUsedFeature)) 'RaiseEvent

        Return theUsedFeature

    End Function 'addPassportFeature


    Private Function _findCustomer(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Customer

        Dim foundCustomer As Customer

        For pLocationFound = 0 To mCustomer.Count - 1
            foundCustomer = _ithCustomer(pLocationFound)

            If foundCustomer IsNot Nothing Then
                If foundCustomer.id = pItemToFind Then
                    Return foundCustomer
                End If
            End If

        Next pLocationFound
        Return Nothing


    End Function '_findItem(pItemToFind,pLocationFound)

    Private Function _findFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Feature

        Dim foundFeature As Feature

        For pLocationFound = 0 To mFeature.Count - 1
            foundFeature = _ithFeature(pLocationFound)

            If foundFeature IsNot Nothing Then
                If foundFeature.id = pItemToFind Then
                    Return foundFeature
                End If
            End If

        Next pLocationFound
        Return Nothing


    End Function '_findItem(pItemToFind,pLocationFound)

    Private Function _findPassportFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            PassportFeature

        Dim foundPassportFeature As PassportFeature

        For pLocationFound = 0 To mPassportFeature.Count - 1
            foundPassportFeature = _ithPassportFeature(pLocationFound)

            If foundPassportFeature IsNot Nothing Then
                If foundPassportFeature.id = pItemToFind Then
                    Return foundPassportFeature
                End If
            End If

        Next pLocationFound
        Return Nothing


    End Function '_findItem(pItemToFind,pLocationFound)

    Private Function _findPassport(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Passport

        Dim foundPassport As Passport

        For pLocationFound = 0 To mPassport.Count - 1
            foundPassport = _ithPassport(pLocationFound)

            If foundPassport IsNot Nothing Then
                If foundPassport.id = pItemToFind Then
                    Return foundPassport
                End If
            End If

        Next pLocationFound
        Return Nothing


    End Function '_findItem(pItemToFind,pLocationFound)

    Private Function _findUsedFeature(
            ByVal pItemToFind As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            UsedFeature

        Dim foundUsedFeature As UsedFeature

        For pLocationFound = 0 To mUsedFeature.Count - 1
            foundUsedFeature = _ithUsedFeature(pLocationFound)

            If foundUsedFeature IsNot Nothing Then
                If foundUsedFeature.id = pItemToFind Then
                    Return foundUsedFeature
                End If
            End If

        Next pLocationFound
        Return Nothing


    End Function '_findItem(pItemToFind,pLocationFound)

    Private Sub _writeFile(ByRef pAppendFile As Boolean)

        Dim outputfile As StreamWriter
        outputfile = New StreamWriter("Transactions-out.txt", append:=pAppendFile)


        ' txtTrans1.Text = "cleared" & vbCrLf & vbCrLf

        For Each item In iterateTransactions()

            outputfile.WriteLine(item)

        Next item
        outputfile.Close()
    End Sub

    Private Sub _readFile()

        Dim inputFile As StreamReader
        ' Dim outputfile As StreamWriter
        Dim errorOutputfile As StreamWriter
        Dim line As String
        Dim leftChar As String
        Dim field() As String
        Dim checker As String
        Dim passbookFeatureChecker As String
        Dim theDate As String
        Dim time As String
        Dim theObject As String
        Dim theMethod As String
        Dim name As String
        Dim customerID As String
        Dim passportID As String
        Dim featureID As String
        Dim passportFeatureID As String
        Dim usedFeatureID As String
        Dim unitMeasurement As String
        Dim quantity As Decimal
        Dim adultPrice As Decimal
        Dim childPrice As Decimal
        Dim purchaseDate As Date
        Dim visitorName As String
        Dim birthdate As Date
        Dim inputValue As String
        Dim dateUsed As Date
        Dim dateUpdated As Date
        Dim locationUsed As String
        Dim foundCustomer As Customer
        Dim foundFeature As Feature
        Dim foundPassport As Passport
        Dim foundPassportFeature As PassportFeature
        Dim foundUsedFeature As UsedFeature
        Dim foundLocation As Integer


        'inputFile = New StreamReader("Data-in-multi.txt")
        inputFile = New StreamReader("Transactions-in.txt")
        'outputfile = New StreamWriter("Data-out-multi.txt")
        'outputfile = New StreamWriter("Data-out-single.txt")
        ' outputfile = New StreamWriter("Transactions-out.txt")
        errorOutputfile = New StreamWriter("Transactions-ERROR.txt", False)




        Do While Not inputFile.EndOfStream

            line = inputFile.ReadLine

            ' txtTrans1.Text &= "temp: Nothing" & vbCrLf
            If line = "" Then
                'outputfile.WriteLine(line)
            Else
                leftChar = line.Substring(0, 1)

                If leftChar = "#" Then
                    'outputfile.WriteLine(line)
                Else

                    field = Split(line, "; ")
                    checker = UCase(Trim(field(2)))
                    passbookFeatureChecker = UCase(Trim(field(3)))

                    Select Case checker

                        Case "CUSTOMER"

                            theDate = UCase(Trim(field(0)))
                            time = UCase(Trim(field(1)))
                            theObject = UCase(Trim(field(2)))
                            theMethod = UCase(Trim(field(3)))
                            customerID = UCase(Trim(field(4)))
                            name = UCase(Trim(field(5)))

                            'outputfile.WriteLine(line)

                            foundCustomer = findCustomer(customerID, foundLocation)

                            If foundCustomer Is Nothing Then 'NOT found
                                addCustomer(customerID, name)
                            Else 'FOUND
                                errorOutputfile.WriteLine(line)
                                errorOutputfile.WriteLine("### Customer ID already exists!" & vbCrLf & vbCrLf)

                            End If


                        Case "FEATURE"

                            theDate = UCase(Trim(field(0)))
                            time = UCase(Trim(field(1)))
                            theObject = UCase(Trim(field(2)))
                            theMethod = UCase(Trim(field(3)))
                            featureID = UCase(Trim(field(4)))
                            name = UCase(Trim(field(5)))
                            unitMeasurement = UCase(Trim(field(6)))
                            adultPrice = Decimal.Parse(UCase(Trim(field(7))))
                            childPrice = Decimal.Parse(UCase(Trim(field(8))))

                            ' outputfile.WriteLine(line)

                            foundFeature = findFeature(featureID, foundLocation)

                            If foundFeature Is Nothing Then 'NOT found
                                If adultPrice >= 0 And childPrice >= 0 Then
                                    addFeature(featureID, name, unitMeasurement, adultPrice, childPrice)
                                Else
                                    errorOutputfile.WriteLine(line)
                                    errorOutputfile.WriteLine("### Price can not be less than 0!" & vbCrLf & vbCrLf)
                                End If

                            Else
                                    errorOutputfile.WriteLine(line)
                                errorOutputfile.WriteLine("### Feature ID already exists!" & vbCrLf & vbCrLf)

                            End If


                        Case "PASSBOOK"

                            Dim format As String = "yyyyMMdd"

                            'parse date from yyyyMMdd format needs provider
                            Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture 'From: https://msdn.microsoft.com/en-us/library/w2sa9yss.aspx 


                            theDate = UCase(Trim(field(0)))
                            time = UCase(Trim(field(1)))
                            theObject = UCase(Trim(field(2)))
                            theMethod = UCase(Trim(field(3)))
                            passportID = UCase(Trim(field(4)))
                            customerID = UCase(Trim(field(5)))
                            purchaseDate = Date.ParseExact(UCase(Trim(field(6))), format, provider)
                            visitorName = UCase(Trim(field(7)))
                            birthdate = Date.ParseExact(UCase(Trim(field(8))), format, provider)

                            ' outputfile.WriteLine(line)


                            inputValue = customerID
                            foundCustomer = findCustomer(inputValue, foundLocation)

                            If foundCustomer Is Nothing Then 'NOT found
                                errorOutputfile.WriteLine(line)
                                errorOutputfile.WriteLine("### Customer not found." & vbCrLf & vbCrLf)

                            Else 'FOUND

                                If birthdate <= Today Then

                                    foundPassport = findPassport(passportID, foundLocation)

                                    If foundPassport Is Nothing Then 'NOT found
                                        addPassport(passportID, purchaseDate, visitorName, birthdate, foundCustomer)
                                    Else
                                        errorOutputfile.WriteLine(line)
                                        errorOutputfile.WriteLine("### Passport ID already exists!" & vbCrLf & vbCrLf)

                                    End If

                                Else
                                    errorOutputfile.WriteLine(line)
                                    errorOutputfile.WriteLine("### This person cant time travel! birthday cant be in the future!" & vbCrLf & vbCrLf)
                                End If


                            End If


                        Case "PASSBOOK_FEATURE"

                            Select Case passbookFeatureChecker  'new case select for passportfeature (purchase, add, update)

                                Case "PURCHASE"

                                    Dim format As String = "yyyyMMdd"

                                    'parse date from yyyyMMdd format needs provider
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture 'From: https://msdn.microsoft.com/en-us/library/w2sa9yss.aspx 


                                    theDate = UCase(Trim(field(0)))
                                    time = UCase(Trim(field(1)))
                                    theObject = UCase(Trim(field(2)))
                                    theMethod = UCase(Trim(field(3)))
                                    passportFeatureID = UCase(Trim(field(4)))
                                    quantity = Decimal.Parse(UCase(Trim(field(5))))
                                    passportID = UCase(Trim(field(6)))
                                    featureID = UCase(Trim(field(7)))

                                    'outputfile.WriteLine(line)

                                    inputValue = passportID
                                    foundPassport = findPassport(inputValue, foundLocation)

                                    If foundPassport Is Nothing Then 'NOT found
                                        errorOutputfile.WriteLine(line)
                                        errorOutputfile.WriteLine("### Passport not found." & vbCrLf & vbCrLf)

                                    Else 'FOUND

                                        inputValue = featureID
                                        foundFeature = findFeature(inputValue, foundLocation)

                                        If foundFeature Is Nothing Then
                                            errorOutputfile.WriteLine(line)
                                            errorOutputfile.WriteLine("### Feature not found." & vbCrLf & vbCrLf)

                                        Else


                                            foundPassportFeature = findPassportFeature(passportFeatureID, foundLocation)

                                            If foundPassportFeature Is Nothing Then 'NOT found
                                                If quantity >= 0 Then

                                                    Dim isChildUnder13 As Boolean
                                                    Dim age As Integer
                                                    isChildUnder13 = False
                                                    Dim price As Decimal

                                                    age = foundPassport.calcAge(foundPassport.Birthdate, Today)
                                                    isChildUnder13 = foundPassport.isChildUnder13(age)
                                                    price = foundPassport.returnPrice(foundPassport, foundFeature, isChildUnder13)

                                                    addPassportFeature(passportFeatureID, quantity, price, foundPassport, foundFeature)
                                                Else
                                                    errorOutputfile.WriteLine(line)
                                                    errorOutputfile.WriteLine("### Quantity can not be less than 0" & vbCrLf & vbCrLf)

                                                End If

                                            Else
                                                    errorOutputfile.WriteLine(line)
                                                errorOutputfile.WriteLine("### Passport Feature ID already exists!" & vbCrLf & vbCrLf)
                                            End If


                                        End If

                                    End If

                                Case "USE"

                                    Dim format As String = "yyyyMMdd"

                                    'parse date from yyyyMMdd format needs provider
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture 'From: https://msdn.microsoft.com/en-us/library/w2sa9yss.aspx 


                                    theDate = UCase(Trim(field(0)))
                                    time = UCase(Trim(field(1)))
                                    theObject = UCase(Trim(field(2)))
                                    theMethod = UCase(Trim(field(3)))
                                    usedFeatureID = UCase(Trim(field(4)))
                                    passportFeatureID = UCase(Trim(field(5)))
                                    dateUsed = Date.ParseExact(UCase(Trim(field(6))), format, provider)
                                    locationUsed = UCase(Trim(field(7)))
                                    quantity = Decimal.Parse(UCase(Trim(field(8))))

                                    'outputfile.WriteLine(line)

                                    inputValue = passportFeatureID
                                    foundPassportFeature = findPassportFeature(inputValue, foundLocation)

                                    If foundPassportFeature Is Nothing Then 'NOT found
                                        errorOutputfile.WriteLine(line)
                                        errorOutputfile.WriteLine("### Passport Feature not Found!" & vbCrLf & vbCrLf)


                                    Else 'FOUND

                                        foundUsedFeature = findUsedFeature(usedFeatureID, foundLocation)


                                        If foundUsedFeature Is Nothing Then 'NOT found
                                            If quantity <= foundPassportFeature.Quantity And quantity >= 0 Then
                                                addUsedFeature(usedFeatureID, dateUsed, locationUsed, quantity, foundPassportFeature)
                                            Else
                                                errorOutputfile.WriteLine(line)
                                                errorOutputfile.WriteLine("### Quantity can not be less than current quantity or less than 0." & vbCrLf & vbCrLf)
                                            End If


                                        Else
                                            errorOutputfile.WriteLine(line)
                                            errorOutputfile.WriteLine("### Used Feature ID already exists!" & vbCrLf & vbCrLf)
                                        End If

                                    End If

                                Case "UPDATE"

                                    Dim format As String = "yyyyMMdd"

                                    'parse date from yyyyMMdd format needs provider
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture 'From: https://msdn.microsoft.com/en-us/library/w2sa9yss.aspx 


                                    theDate = UCase(Trim(field(0)))
                                    time = UCase(Trim(field(1)))
                                    theObject = UCase(Trim(field(2)))
                                    theMethod = UCase(Trim(field(3)))
                                    passportFeatureID = UCase(Trim(field(4)))
                                    dateUpdated = Date.ParseExact(UCase(Trim(field(5))), format, provider)
                                    quantity = Decimal.Parse(UCase(Trim(field(6))))

                                    'outputfile.WriteLine(line)

                                    inputValue = passportFeatureID
                                    foundPassportFeature = findPassportFeature(inputValue, foundLocation)

                                    If foundPassportFeature Is Nothing Then 'NOT found
                                        errorOutputfile.WriteLine(line)
                                        errorOutputfile.WriteLine("### Passport Feature not found" & vbCrLf & vbCrLf)
                                    Else 'FOUND

                                        If quantity >= 0 Then
                                            updatePassportFeature(passportFeatureID, quantity, foundPassportFeature.Price, foundPassportFeature, foundPassportFeature.Passport, foundPassportFeature.Feature, dateUpdated)
                                        Else
                                            errorOutputfile.WriteLine(line)
                                            errorOutputfile.WriteLine("### Updated quantity can not be less than 0." & vbCrLf & vbCrLf)
                                        End If


                                    End If




                            End Select 'end select passbookfeature type


                    End Select 'end select action type


                End If 'if block that it is not comment line
            End If 'if block that it is not blank

        Loop



        inputFile.Close()
        errorOutputfile.Close()

    End Sub

    Private Function _averageBalanceUnused() As String

        Dim unused As Decimal
        Dim totalUnused As Decimal = 0
        Dim sumPassportFeatures As Decimal = 0

        For Each item As PassportFeature In iteratePassportFeature()

            unused = item.Quantity * item.Price
            totalUnused += unused
            sumPassportFeatures += item.Quantity

        Next item

        If totalUnused > 0 Then

            ' txtMetricsTabSummaryInfoTbcMain.Text &= "totalUnused: " & totalUnused & vbCrLf
            ' txtMetricsTabSummaryInfoTbcMain.Text &= "sumPassportFeatures: " & sumPassportFeatures & vbCrLf
            Return "Average balance of unused PassbookFeatures In Dollars: " & "$" & (totalUnused / sumPassportFeatures).ToString("N2") & vbCrLf & vbCrLf

        Else

            Return "Average balance of unused PassbookFeatures In Dollars: " & "N/A" & vbCrLf & vbCrLf

        End If

    End Function

    Private Function _sumUnusedPassportFeature() As String

        Dim unused As Decimal
        Dim totalUnused As Decimal = 0


        For Each item As PassportFeature In iteratePassportFeature()

            unused = item.Quantity * item.Price
            totalUnused += unused


        Next item

        Return "Sum of unused Passport Features In Dollars: $" & totalUnused.ToString("N2") & vbCrLf & vbCrLf



    End Function

    Private Function _mostPopularPassportFeature() As String

        Dim count = NumberFeatures - 1
        Dim Features(count) As Decimal
        Dim numberOfFeatures As Integer = 0
        Dim currentHighestIndex As Integer = 0


        For Each item As Feature In iterateFeature()


            For Each passportFeature As PassportFeature In iteratePassportFeature()

                If passportFeature.Feature.id = item.id Then
                    Features(numberOfFeatures) += passportFeature.Quantity
                End If

            Next

            numberOfFeatures += 1
        Next

        For i = 0 To count
            If Features(i) > Features(currentHighestIndex) Then
                currentHighestIndex = i
            End If
        Next

        If _ithFeature(currentHighestIndex) Is Nothing Then
            Return Nothing
        Else

            Return "Most Popular Passport Feature: " & _ithFeature(currentHighestIndex).id.ToString & vbCrLf & vbCrLf

        End If
    End Function


    Private Function _averagePassportsPerCustomer() As String


        'Average number of Passbooks per Customer

        Dim numberCustomers As Decimal = 0
        Dim numberPassports As Decimal = 0

        For Each item As Customer In iterateCustomer()

            numberCustomers += 1

        Next item

        For Each item As Passport In iteratePassport()

            numberPassports += 1

        Next item

        If numberPassports > 0 Then

            ' txtMetricsTabSummaryInfoTbcMain.Text &= "numberCustomers: " & numberCustomers & vbCrLf
            'txtMetricsTabSummaryInfoTbcMain.Text &= "numberPassports: " & numberPassports & vbCrLf
            Return "Average number of Passbooks per Customer: " & (numberPassports / numberCustomers).ToString("N2") & vbCrLf & vbCrLf
        Else

            Return "Average balance of unused PassbookFeatures In Dollars: " & "N/A" & vbCrLf & vbCrLf

        End If


    End Function

    Private Function _percentPassportFeaturesUsed() As String

        Dim FeaturesDollar As Decimal
        Dim totalFeaturesDollar As Decimal = 0

        Dim usedFeaturesDollar As Decimal
        Dim usedTotalFeaturesDollar As Decimal = 0


        For Each item As PassportFeature In iteratePassportFeature()

            FeaturesDollar = item.Quantity * item.Price
            totalFeaturesDollar += FeaturesDollar

        Next item


        For Each item As UsedFeature In iterateUsedFeature()

            usedFeaturesDollar = item.QuantityUsed * item.PassportFeature.Price
            usedTotalFeaturesDollar += usedFeaturesDollar

        Next item

        If totalFeaturesDollar > 0 Then
            Return "Percent of Passbook Features Used: " & (usedTotalFeaturesDollar / totalFeaturesDollar).ToString("N2") & ("%") & vbCrLf & vbCrLf

        Else
            Return "Percent of Passbook Features Used: N/A" & vbCrLf & vbCrLf
        End If



    End Function

    Private Function _averageAge() As String

        Dim totalAge As Decimal = 0
        Dim passbookHolders As Decimal = 0

        For Each item As Passport In iteratePassport()

            totalAge += item.calcAge(item.Birthdate, item.DatePurchased)
            passbookHolders += 1

        Next item

        If passbookHolders > 0 Then
            Return "Average Age of Passport Holders: " & (totalAge / passbookHolders).ToString("N2") & vbCrLf & vbCrLf

        Else
            Return "Average Age of Passport Holders: N/A" & vbCrLf & vbCrLf
        End If

    End Function

    Private Function _passportHoldersBirthday() As String

        Dim count As Integer = 0

        For Each item As Passport In iteratePassport()

            If item.Birthdate.Month = Now.Month Then
                count += 1
            End If

        Next item

        Return "Passport Holders with Birthday this Month: " & count.ToString

    End Function



    Private Function _toString() As String

        Dim tmpStr As String

        tmpStr = "ThemePark =" & _ThemeParkName _
            & " #Customers = " & _NumberCustomers _
            & ", #Passports = " & _NumberPassports.ToString _
            & ", #Features = " & _NumberFeatures.ToString _
            & ", #PassportFeatures = " & _NumberPassportFeatures.ToString _
            & ", #UsedFeatures = " & _NumberUsedFeatures.ToString _
            & vbCrLf



        Return tmpStr

    End Function '_toString()



#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined.
    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system


    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running

#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'No Events are currently defined.
    'These are all public.

    Public Event ThemePark_CustomerAdded(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )

    Public Event ThemePark_PassportFeatureAdded(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )

    Public Event ThemePark_FeatureAdded(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )

    Public Event ThemePark_PassportAdded(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )

    Public Event ThemePark_UsedFeature(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )

    Public Event ThemePark_UpdatePassportFeature(
    ByVal sender As System.Object,
    ByVal e As System.EventArgs
    )


#End Region 'Events

End Class 'ThemePark