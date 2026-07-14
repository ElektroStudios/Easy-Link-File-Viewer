
#Region " Usage Examples "

' <TypeConverter(GetType(EnumDescriptionConverter))>
' <Flags>
' Public Enum TestEnum
'     <Description("First")> One = 1
'     <Description("Second")> Two = 2
'     <Description("Third")> Three = 3
' End Enum
'
' Public Class TestClass
'     <DefaultValue(TestEnum.MyUpperCamelCaseName)>
'     Public Property TestProperty As TestEnum = TestEnum.MyUpperCamelCaseName
' End Class
'
' Public Class Form1
'     Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
'         Me.PropertyGrid.SelectedObject = New TestClass()
'     End Sub
' End Class

#End Region

#Region " Option Statements "

Option Strict Off
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.ComponentModel
Imports System.Globalization
Imports System.Linq
Imports System.Reflection

#End Region

#Region " EnumDescriptionConverter "

' ReSharper disable once CheckNamespace

Namespace DevCase.Runtime.TypeConverters

    ''' <summary>
    ''' Provides a way to convert <see cref="[Enum]"/> fields and flag combinations to and from their <see cref="DescriptionAttribute"/> attributes.
    ''' <para></para>
    ''' Note: A enumeration field must have a <see cref="DescriptionAttribute"/> attribute defined.
    ''' </summary>
    ''' 
    ''' <example> This is a code example.
    ''' <code language="VB">
    ''' &lt;TypeConverter(GetType(EnumDescriptionConverter))&gt;
    ''' &lt;Flags&gt;
    ''' Public Enum TestEnum
    '''     &lt;Description("First")&gt; One = 1
    '''     &lt;Description("Second")&gt; Two = 2
    '''     &lt;Description("Third")&gt; Three = 3
    ''' End Enum
    '''
    ''' Public Class TestClass
    '''     &lt;DefaultValue(TestEnum.MyUpperCamelCaseName)&gt;
    '''     Public Property TestProperty As TestEnum = TestEnum.MyUpperCamelCaseName
    ''' End Class
    '''
    ''' Public Class Form1
    '''     Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
    '''         Me.PropertyGrid.SelectedObject = New TestClass()
    '''     End Sub
    ''' End Class
    ''' </code>
    ''' </example>
    ''' 
    ''' <seealso cref="EnumConverter"/>
    Public NotInheritable Class EnumDescriptionConverter : Inherits EnumConverter

#Region " Private Fields "

        Private ReadOnly Property HasFlags As Boolean = Attribute.IsDefined(Me.EnumType, GetType(FlagsAttribute))

#End Region

#Region " Constructors "

        ''' <summary>
        ''' Initializes a new instance of the <see cref="EnumDescriptionConverter"/> class.
        ''' </summary>
        ''' 
        ''' <param name="type">
        ''' A <see cref="Type"/> that represents the type of enumeration to associate with this enumeration converter.
        ''' </param>
        Public Sub New(type As Type)
            MyBase.New(type)
        End Sub

#End Region

#Region " Public Methods "

        ''' <summary>
        ''' Gets a value indicating whether this object supports a standard set of 
        ''' values that can be picked from a list using the specified context.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext" /> that provides a format context.
        ''' </param>
        ''' 
        ''' <returns>
        ''' Returns <see langword="True"/> because <see cref="TypeConverter.GetStandardValues"/> 
        ''' should be called to find a common set of values the object supports. 
        ''' <para></para>
        ''' This method never returns <see langword="False"/>.
        ''' </returns>
        Public Overrides Function GetStandardValuesSupported(context As ITypeDescriptorContext) As Boolean

            Return True
        End Function

        ''' <summary>
        ''' Gets a value indicating whether the list of standard values returned from 
        ''' <see cref="TypeConverter.GetStandardValues"/> is an exclusive list using the specified context.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <returns>
        ''' Returns <see langword="True"/> if the <see cref="StandardValuesCollection"/> returned from 
        ''' <see cref="TypeConverter.GetStandardValues"/> is an exhaustive list of possible values; 
        ''' Returns <see langword="False"/> if other values are possible.
        ''' </returns>
        Public Overrides Function GetStandardValuesExclusive(context As ITypeDescriptorContext) As Boolean

            Return False
        End Function

        ''' <summary>
        ''' Gets a collection of standard values for the data type this validator is designed for.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext" /> that provides a format context.
        ''' </param>
        ''' 
        ''' <returns>
        ''' A <see cref="StandardValuesCollection"/> that holds a standard set of valid values, 
        ''' or <see langword="null"/> if the data type does not support a standard set of values.
        ''' </returns>
        Public Overrides Function GetStandardValues(context As ITypeDescriptorContext) As StandardValuesCollection

            Dim values As New SortedSet(Of Object)

            For Each field As FieldInfo In Me.EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
                Dim value As Object = field.GetValue(Nothing)
                values.Add(value)
            Next

            Return New StandardValuesCollection(values)
        End Function

        ''' <summary>
        ''' Returns whether this converter can convert the object to the specified type, using the specified context.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="destinationType">
        ''' A <see cref="Type"/> that represents the type you want to convert to.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if this converter can perform the conversion; otherwise, <see langword="False"/>.
        ''' </returns>
        Public Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean

            Return destinationType Is GetType(String) OrElse
               MyBase.CanConvertTo(context, destinationType)

        End Function

        ''' <summary>
        ''' Returns whether this converter can convert an object of the given type to the type of this converter, 
        ''' using the specified context.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="sourceType">
        ''' A <see cref="Type"/> that represents the type you want to convert from.
        ''' </param>
        ''' 
        ''' <returns>
        ''' <see langword="True"/> if this converter can perform the conversion; otherwise, <see langword="False"/>.
        ''' </returns>
        Public Overrides Function CanConvertFrom(context As ITypeDescriptorContext, sourceType As Type) As Boolean

            Return sourceType Is GetType(String) OrElse
               MyBase.CanConvertFrom(context, sourceType)

        End Function

        ''' <summary>
        ''' Converts the given value object to the specified type, using the specified context and culture information.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="culture">
        ''' A <see cref="CultureInfo"/>. If null is passed, the current culture is assumed.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The <see cref="Object"/> to convert.
        ''' </param>
        ''' 
        ''' <param name="destinationType">
        ''' The <see cref="Type"/> to convert the <paramref name="value"/> parameter to.
        ''' </param>
        ''' 
        ''' <returns>
        ''' An <see cref="Object"/> that represents the converted value.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As CultureInfo, value As Object, destinationType As Type) As Object

            If destinationType Is GetType(String) Then

                If Not Me.HasFlags Then
                    Dim fi As FieldInfo = Me.EnumType.GetField([Enum].GetName(Me.EnumType, value))
                    Dim descAttr As DescriptionAttribute =
                    CType(Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)
                    Return If(descAttr IsNot Nothing, descAttr.Description, value.ToString())

                Else
                    ' Allows to parse values like "None" (zero) or "All".
                    Dim exactField As FieldInfo =
                    Me.EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static).
                                FirstOrDefault(Function(fi As FieldInfo) Object.Equals(fi.GetValue(Nothing), value))
                    If exactField IsNot Nothing Then
                        Dim exactDescAttr As DescriptionAttribute =
                        CType(Attribute.GetCustomAttribute(exactField, GetType(DescriptionAttribute)), DescriptionAttribute)
                        If exactDescAttr IsNot Nothing Then
                            Dim desc As String = exactDescAttr.Description
                            If Not desc.Contains(","c) Then
                                Return desc
                            Else
                                'Throw New Exception($"Description contains a comma: {desc}")
                                Return exactField.Name
                            End If
                        Else
                            Return exactField.Name
                        End If

                    Else
                        Dim names As New List(Of String)

                        Dim fieldsOrdered As FieldInfo() =
                            Me.EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static).
                                        OrderBy(Function(f As FieldInfo) f.GetValue(Nothing)).ToArray()

                        For Each fi As FieldInfo In fieldsOrdered
                            Dim fiValue As Object = fi.GetValue(Nothing)
                            If (value And fiValue) = fiValue AndAlso fiValue <> 0 Then
                                Dim descAttr As DescriptionAttribute =
                                CType(Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)
                                If descAttr IsNot Nothing Then
                                    Dim desc As String = descAttr.Description
                                    ' If 'DescriptionAttribute' contains a comma, adds the field name instead.
                                    ' This prevents user mistakes.
                                    If Not desc.Contains(","c) Then
                                        names.Add(desc)
                                    Else
                                        ' Throw New Exception($"Description contains a comma: {desc}")
                                        names.Add(fi.Name)
                                    End If
                                Else
                                    names.Add(fi.Name)
                                End If

                            End If
                        Next

                        Return String.Join(", ", names)
                    End If
                End If
            Else

                Return MyBase.ConvertTo(context, culture, value, destinationType)
            End If

        End Function

        ''' <summary>
        ''' Converts the given object to the type of this converter, using the specified context and culture information.
        ''' </summary>
        ''' 
        ''' <param name="context">
        ''' An <see cref="ITypeDescriptorContext"/> that provides a format context.
        ''' </param>
        ''' 
        ''' <param name="culture">
        ''' The <see cref="CultureInfo"/> to use as the current culture.
        ''' </param>
        ''' 
        ''' <param name="value">
        ''' The <see cref="Object"/> to convert.
        ''' </param>
        ''' 
        ''' <returns>
        ''' An <see cref="Object"/> that represents the converted value.
        ''' </returns>
        <DebuggerStepThrough>
        Public Overrides Function ConvertFrom(context As ITypeDescriptorContext, culture As CultureInfo, value As Object) As Object

            If Not Me.HasFlags Then
                For Each fi As FieldInfo In Me.EnumType.GetFields()
                    Dim descAttr As DescriptionAttribute =
                        CType(Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)
                    If (descAttr IsNot Nothing) AndAlso DirectCast(value, String) = descAttr.Description Then
                        Return [Enum].Parse(Me.EnumType, fi.Name, ignoreCase:=False)
                    End If
                Next fi
                Return [Enum].Parse(Me.EnumType, DirectCast(value, String))

            Else
                ' Allows to parse a numeric value (e.g. "3" > "Flag 1, Flag 2")
                Dim numericValue As ULong
                If ULong.TryParse(DirectCast(value, String).Trim(), numericValue) Then

                    Dim combinedValue As ULong = 0
                    For Each fi As FieldInfo In Me.EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
                        Dim val As Object = fi.GetValue(Nothing)
                        If val <> 0 Then combinedValue = combinedValue Or val
                    Next

                    If (numericValue And Not combinedValue) = 0 Then
                        Return [Enum].ToObject(Me.EnumType, numericValue)
                    Else
                        Throw New ArgumentException($"The value '{numericValue}' is not a valid combination of defined flags in '{Me.EnumType.Name}' enumeration.")
                    End If
                End If

                Dim result As Object = 0
                Dim names As String() = DirectCast(value, String).Split(","c).Select(Function(p) p.Trim()).ToArray()
                For Each name As String In names
                    Dim matched As Boolean = False
                    For Each fi As FieldInfo In Me.EnumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
                        Dim descAttr As DescriptionAttribute =
                        CType(Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)
                        If (descAttr IsNot Nothing AndAlso String.Equals(descAttr.Description, name, StringComparison.OrdinalIgnoreCase)) OrElse
                       String.Equals(fi.Name, name, StringComparison.OrdinalIgnoreCase) Then
                            result = result Or fi.GetValue(Nothing)
                            matched = True
                            Exit For
                        End If
                    Next

                    If Not matched Then
                        Throw New ArgumentException($"No entry was found for '{name}' in the enumeration '{Me.EnumType.Name}'.")
                    End If
                Next

                Return [Enum].ToObject(Me.EnumType, result)
            End If

        End Function

#End Region

    End Class

End Namespace

#End Region
