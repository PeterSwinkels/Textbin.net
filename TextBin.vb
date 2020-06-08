'The imports and settings used by this class.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.VisualBasic.ControlChars
Imports System
Imports System.Convert
Imports System.IO
Imports System.Text

'This class contains the Text Bin interface.
Public Class TextBinClass
   Public Event FoundText(Text As String, ByRef ContinueSearch As Boolean) 'The event raised when a character is found that is not considered to be text.
   Public Event HandleError(Exceptiono As Exception)                       'The event raised when an error occurs.

   'This enumeration contains the categories of characters used by this class.
   Private Enum CharacterCategoriesE As Integer
      Unreadable = 0         'The character is not "human readable."
      RangeCharacter         'The character falls inside the specified range.
      AdditionalCharacter    'The character is a "human readable" character outside the specified range.
      ExcludedCharacter      'The character is excluded from the "human readable" characters.
      UnicodeNullCharacter   'The character is null character between two "human readable" characters.
   End Enum

   'This structure  defines what is considered to be text.
   Private Structure TextDefinitionStr
      Public RangeStart As Integer       'Defines the first character in the human readable character range.
      Public RangeEnd As Integer         'Defines the last character in the human readable character range.
      Public Additional As String        'Defines any characters outside the defined range, but should be included.
      Public Excluded As String          'Defines any characters inside the defined range, but should be excluded.
      Public IncludeUnicode As Boolean   'Indicates that single null characters between two "text" characters are ignored.
   End Structure

   'The variables used by this class:
   Private TextDefinition As New TextDefinitionStr With {.Additional = Cr & Tab, .Excluded = "", .IncludeUnicode = True, .RangeEnd = ToByte("~"c), .RangeStart = ToByte(" "c)} 'Contains the definition of what is considered to be text.

   'This procedure returns the specified character's category.
   Private Function CharacterCategory(Character As Integer, Optional PreviousCharacter As Integer = Nothing, Optional NextCharacter As Integer = Nothing, Optional NextNextCharacter As Integer = Nothing) As CharacterCategoriesE
      Try
         Dim Category As CharacterCategoriesE = CharacterCategoriesE.Unreadable

         With TextDefinition
            If .Excluded.Contains(ToChar(Character)) Then
               Category = CharacterCategoriesE.ExcludedCharacter
            ElseIf ToChar(Character) = NullChar AndAlso ToChar(NextNextCharacter) = NullChar Then
               If .IncludeUnicode Then
                  If Not (PreviousCharacter = Nothing OrElse NextCharacter = Nothing) Then
                     If Not (.Excluded.Contains(ToChar(PreviousCharacter)) OrElse .Excluded.Contains(ToChar(NextCharacter))) Then
                        If (PreviousCharacter >= .RangeStart AndAlso PreviousCharacter <= .RangeEnd) OrElse .Additional.Contains(ToChar(PreviousCharacter)) Then
                           If (NextCharacter >= .RangeStart AndAlso NextCharacter <= .RangeEnd) OrElse .Additional.Contains(ToChar(NextCharacter)) Then
                              Category = CharacterCategoriesE.UnicodeNullCharacter
                           End If
                        End If
                     End If
                  End If
               End If
            ElseIf Character >= .RangeStart AndAlso Character <= .RangeEnd Then
               Category = CharacterCategoriesE.RangeCharacter
            ElseIf .Additional.Contains(ToChar(Character)) Then
               Category = CharacterCategoriesE.AdditionalCharacter
            End If
         End With

         Return Category
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return CharacterCategoriesE.Unreadable
   End Function

   'This procedure changes the text definition used by this class.
   Public Sub DefineText(RangeStart As Integer, RangeEnd As Integer, Optional Additional As String = "", Optional Excluded As String = "", Optional IncludeUnicode As Boolean = False)
      Try
         With TextDefinition
            .RangeStart = RangeStart
            .RangeEnd = RangeEnd
            .Additional = Additional
            .Excluded = Excluded
            .IncludeUnicode = IncludeUnicode
         End With
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure searches the binary data for strings of human readable characters.
   Public Sub FindText(BinaryData() As Byte)
      Try
         Dim Category As CharacterCategoriesE = CharacterCategoriesE.Unreadable
         Dim Character As Integer = Nothing
         Dim ContinueSearch As Boolean = False
         Dim NextCharacter As Integer = Nothing
         Dim NextNextCharacter As Integer = Nothing
         Dim PreviousCategory As CharacterCategoriesE = CharacterCategoriesE.Unreadable
         Dim PreviousCharacter As Integer = Nothing
         Dim Text As New StringBuilder

         For Index As Integer = BinaryData.GetLowerBound(0) To BinaryData.GetUpperBound(0)
            PreviousCharacter = Character
            Character = BinaryData(Index)
            NextCharacter = If(Index + 1 <= BinaryData.GetUpperBound(0), BinaryData(Index + 1), Nothing)
            NextNextCharacter = If(Index + 2 <= BinaryData.GetUpperBound(0), BinaryData(Index + 2), Nothing)

            PreviousCategory = Category
            Category = CharacterCategory(Character, PreviousCharacter, NextCharacter, NextNextCharacter)

            If Category = CharacterCategoriesE.UnicodeNullCharacter Then
               If Not (PreviousCategory = CharacterCategoriesE.AdditionalCharacter OrElse PreviousCategory = CharacterCategoriesE.RangeCharacter) Then
                  Category = CharacterCategoriesE.Unreadable
               End If
            ElseIf Not Category = CharacterCategoriesE.UnicodeNullCharacter Then
               If Category = CharacterCategoriesE.AdditionalCharacter OrElse Category = CharacterCategoriesE.RangeCharacter Then
                  Text.Append(ToChar(Character))
               ElseIf Category = CharacterCategoriesE.ExcludedCharacter OrElse Category = CharacterCategoriesE.Unreadable Then
                  If Text.Length > 0 Then
                     RaiseEvent FoundText(Text.ToString(), ContinueSearch)
                     If Not ContinueSearch Then Exit Sub
                     Text.Clear()
                  End If
               End If
            End If
         Next Index
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try
   End Sub

   'This function returns the binary data from the specified file.
   Public Function GetBinaryData(BinaryFile As String) As Byte()
      Try
         If Not BinaryFile = Nothing Then Return File.ReadAllBytes(BinaryFile)
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try

      Return {}
   End Function

   'This procedure returns the text definition used by this class.
   Public Sub GetTextDefinition(Optional ByRef RangeStart As Integer = Nothing, Optional ByRef RangeEnd As Integer = Nothing, Optional ByRef Additional As String = "", Optional ByRef Excluded As String = "", Optional ByRef IncludeUnicode As Boolean = False)
      Try
         With TextDefinition
            RangeStart = .RangeStart
            RangeEnd = .RangeEnd
            Additional = .Additional
            Excluded = .Excluded
            IncludeUnicode = .IncludeUnicode
         End With
      Catch ExceptionO As Exception
         RaiseEvent HandleError(ExceptionO)
      End Try
   End Sub
End Class
