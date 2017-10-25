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

Public Class UsedFeature

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables
    Private mId As String
    Private mDateUsed As Date
    Private mLocationWhereUsed As String
    Private mQuantityUsed As Decimal
    Private mPassportFeature As PassportFeature

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
            ByVal pDateUsed As Date,
            ByVal pLocationWhereUsed As String,
            ByVal pQuantityUsed As Decimal,
            ByVal pPassportFeature As PassportFeature
            )

        MyBase.New()

        _id = pId
        _DateUsed = pDateUsed
        _LocationWhereUsed = pLocationWhereUsed
        _QuantityUsed = pQuantityUsed
        _PassportFeature = pPassportFeature

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

    Public Property id As String
        Get
            Return _id()
        End Get
        Set(pValue As String)
            _id() = pValue
        End Set
    End Property

    Public Property PassportFeature As PassportFeature
        Get
            Return _PassportFeature()
        End Get
        Set(pValue As PassportFeature)
            _PassportFeature() = pValue
        End Set
    End Property

    Public Property DateUsed As Date
        Get
            Return _DateUsed()
        End Get
        Set(pValue As Date)
            _DateUsed() = pValue
        End Set
    End Property

    Public Property LocationWhereUsed As String
        Get
            Return _LocationWhereUsed()
        End Get
        Set(pValue As String)
            _LocationWhereUsed() = pValue
        End Set
    End Property

    Public Property QuantityUsed As Decimal
        Get
            Return _QuantityUsed()
        End Get
        Set(pValue As Decimal)
            _QuantityUsed() = pValue
        End Set
    End Property



    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _id() As String
        Get
            Return mId
        End Get
        Set(ByVal pValue As String)
            mId = pValue
        End Set
    End Property

    Private Property _PassportFeature() As PassportFeature
        Get
            Return mPassportFeature
        End Get
        Set(ByVal pValue As PassportFeature)
            mPassportFeature = pValue
        End Set
    End Property

    Private Property _DateUsed() As Date
        Get
            Return mDateUsed
        End Get
        Set(ByVal pValue As Date)
            mDateUsed = pValue
        End Set
    End Property

    Private Property _LocationWhereUsed() As String
        Get
            Return mLocationWhereUsed
        End Get
        Set(ByVal pValue As String)
            mLocationWhereUsed = pValue
        End Set
    End Property

    Private Property _QuantityUsed() As Decimal
        Get
            Return mQuantityUsed
        End Get
        Set(ByVal pValue As Decimal)
            mQuantityUsed = pValue
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

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _toString() As String

        Dim tmpStr As String

        tmpStr = vbCrLf _
            & "ID: " & mId & vbCrLf _
            & "Date Used: " & mDateUsed & vbCrLf _
            & "Location Where Used: " & mLocationWhereUsed & vbCrLf _
            & "Quantity Used: " & mQuantityUsed & vbCrLf & vbCrLf _
            & "(PASSED OBJECT) Passport Feature: " & mPassportFeature.ToString

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

End Class 'UsedFeature