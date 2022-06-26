Imports System.Data
#If NETFRAMEWORK Then
Imports System.Data.SqlClient
#Else
Imports Microsoft.Data.SqlClient
#End If

Module modPigSQLSrvLib

    Public Function GetDataCategoryBySqlDbType(SqlDbType As SqlDbType) As Field.EnumDataCategory
        Try
            Select Case SqlDbType
                Case SqlDbType.Char, SqlDbType.VarChar, SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.NText, SqlDbType.Xml
                    Return Field.EnumDataCategory.StrValue
                Case SqlDbType.BigInt, SqlDbType.TinyInt, SqlDbType.SmallInt, SqlDbType.Int
                    Return Field.EnumDataCategory.IntValue
                Case SqlDbType.Decimal, SqlDbType.Real, SqlDbType.Float, SqlDbType.Money, SqlDbType.SmallMoney
                    Return Field.EnumDataCategory.DecValue
                Case SqlDbType.DateTime, SqlDbType.Date, SqlDbType.DateTime2, SqlDbType.SmallDateTime
                    Return Field.EnumDataCategory.DateValue
                Case SqlDbType.Bit
                    Return Field.EnumDataCategory.BooleanValue
                Case Else
                    Return Field.EnumDataCategory.OtherValue
            End Select
        Catch ex As Exception
            Return Field.EnumDataCategory.OtherValue
        End Try
	End Function

End Module
