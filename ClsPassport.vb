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
#End Region 'Option / Imports

Public Class Passport

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mId As String
    Private mDatePurchased As Date
    Private mVisitorName As String
    Private mBirthdate As Date
    ' Private mAge As Integer
    'Private mIsChild As Boolean
    Private mOwner As Customer


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
            ByVal pId As String,
            ByVal pDatePurchased As Date,
            ByVal pVisitorName As String,
            ByVal pBirthdate As Date,
            ByVal pOwner As Customer
            )

        MyBase.New()

        _Id = pId
        _DatePurchased = pDatePurchased
        _VisitorName = pVisitorName
        _Birthdate = pBirthdate
        _Owner = pOwner
        '_Age = _calcAge(pBirthdate, pDatePurchased)
        '_IsChild = _isChildUnder13(pBirthdate, pDatePurchased)

    End Sub 'New

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    'Public ReadOnly Property id As String
    '    Get
    '        Return _Id()
    '    End Get
    '    'Set(pValue As String)
    '    '    _id() = pValue
    '    'End Set
    'End Property

    Public Property id As String
        Get
            Return _Id()
        End Get
        Set(pValue As String)
            _Id() = pValue
        End Set
    End Property

    Public Property Owner As Customer
        Get
            Return _Owner()
        End Get
        Set(pValue As Customer)
            _Owner() = pValue
        End Set
    End Property

    Public Property DatePurchased As Date
        Get
            Return _DatePurchased()
        End Get
        Set(pValue As Date)
            _DatePurchased() = pValue
        End Set
    End Property

    Public Property VisitorName As String
        Get
            Return _VisitorName()
        End Get
        Set(pValue As String)
            _VisitorName() = pValue
        End Set
    End Property

    Public Property Birthdate As Date
        Get
            Return _Birthdate()
        End Get
        Set(pValue As Date)
            _Birthdate() = pValue
        End Set
    End Property

    'Public Property Age As Integer
    '    Get
    '        Return _Age()
    '    End Get
    '    Set(pValue As Integer)
    '        _Age() = pValue
    '    End Set
    'End Property

    'Public Property IsChild As Boolean
    '    Get
    '        Return _IsChild()
    '    End Get
    '    Set(pValue As Boolean)
    '        _IsChild() = pValue
    '    End Set
    'End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _Id() As String
        Get
            Return mId
        End Get
        Set(ByVal pValue As String)
            mId = pValue
        End Set
    End Property

    Private Property _Owner() As Customer
        Get
            Return mOwner
        End Get
        Set(ByVal pValue As Customer)
            mOwner = pValue
        End Set
    End Property

    Private Property _DatePurchased() As Date
        Get
            Return mDatePurchased
        End Get
        Set(ByVal pValue As Date)
            mDatePurchased = pValue
        End Set
    End Property

    Private Property _VisitorName() As String
        Get
            Return mVisitorName
        End Get
        Set(ByVal pValue As String)
            mVisitorName = pValue
        End Set
    End Property

    Private Property _Birthdate() As Date
        Get
            Return mBirthdate
        End Get
        Set(ByVal pValue As Date)
            mBirthdate = pValue
        End Set
    End Property

    'Private Property _Age() As Integer
    '    Get
    '        Return mAge
    '    End Get
    '    Set(ByVal pValue As Integer)
    '        mAge = pValue
    '    End Set
    'End Property

    'Private Property _IsChild() As Boolean
    '    Get
    '        Return mIsChild
    '    End Get
    '    Set(ByVal pValue As Boolean)
    '        mIsChild = pValue
    '    End Set
    'End Property


#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Function returnPrice(ByVal pPassport As Passport, ByVal pFeature As Feature, ByVal pIsChildUnder13 As Boolean) As Decimal

        Return _returnPrice(pPassport, pFeature, pIsChildUnder13)

    End Function 'returnPrice

    Public Function calcAge(ByVal pBirthDate As Date, ByVal pDateAsOf As Date) As Integer

        Return _calcAge(pBirthDate, pDateAsOf)

    End Function 'calcAge
    Public Function isChildUnder13(ByVal pAge As Integer) As Boolean

        Return _isChildUnder13(pAge)

    End Function 'isChildUnder13



    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _returnPrice(ByVal pPassport As Passport, ByVal pFeature As Feature, ByVal pIsChildUnder13 As Boolean) As Decimal
        'Returns price based on whether passed in passport is an adult or child

        If (pIsChildUnder13) Then

            Return pFeature.ChildPricePerUnit

        End If

        Return pFeature.AdultPricePerUnit

    End Function '_returnPrice


    Private Function _isChildUnder13(ByVal pAge As Integer) As Boolean
        'Shared function only used to be able to pass data from process data. We do not have objects for passports before we need to find out how old they are.
        'Might be edited before project 4

        Dim isChildUnder13 As Boolean

        isChildUnder13 = False


        If pAge < 13 Then
            isChildUnder13 = True
        End If

        Return isChildUnder13

    End Function '_isChildUnder13

    Private Function _calcAge(ByVal pBirthDate As Date, ByVal pDateAsOf As Date) As Integer
        'Shared function only used to be able to pass data from process data. We do not have objects for passports before we need to find out if they are a child.
        'Might be edited before project 4

        'Dim age As Double
        'Dim age2 As Integer

        'age = pDateAsOf.Subtract(pBirthDate).TotalDays / 365.25
        ''Formula assitance credited to Randy on http://codereview.stackexchange.com/questions/10263/attempt-to-calculate-age-in-vb-net
        'age = Math.Truncate(age)
        'age2 = Convert.ToInt32(age)

        'Return age2

        Dim age As Long
        Dim age2 As Integer

        age = DateDiff("yyyy", pBirthDate, pDateAsOf)

        If pDateAsOf < DateSerial(Year(pDateAsOf), Month(pBirthDate), Weekday(pBirthDate)) Then
            age = age - 1
        End If
        'Formula assistance to microsoft at https://msdn.microsoft.com/en-us/enus/library/aa227466(v=vs.60).aspx 

        age2 = Convert.ToInt16(age)
        Return age2

    End Function '_calcAge


    Private Function _toString() As String

        Dim tmpStr As String

        tmpStr = vbCrLf & "ID: " & mId & vbCrLf _
            & "Date Purchased: " & mDatePurchased & vbCrLf _
            & "Visitor Name: " & mVisitorName & vbCrLf _
            & "Birthdate: " & mBirthdate & vbCrLf _
            & "Age: " & calcAge(mBirthdate, mDatePurchased).ToString & vbCrLf _
            & "Is Child? " & isChildUnder13(calcAge(mBirthdate, mDatePurchased)).ToString & vbCrLf & vbCrLf _
          & "(PASSED OBJECT) Owner: " & mOwner.ToString



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

#End Region 'Events

End Class 'Passport