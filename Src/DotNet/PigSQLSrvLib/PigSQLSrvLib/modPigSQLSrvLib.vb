Imports System.Data
Imports System.Data.SqlClient

Module modPigSQLSrvLib

	Public Function GetDataCategoryBySqlDbType(SqlDbType As SqlDbType) As Field.DataCategoryEnum
		Try
			Select Case SqlDbType
				Case SqlDbType.Char, SqlDbType.VarChar, SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.NText, SqlDbType.Xml
					Return Field.DataCategoryEnum.StrValue
				Case SqlDbType.BigInt, SqlDbType.TinyInt, SqlDbType.SmallInt, SqlDbType.Int
					Return Field.DataCategoryEnum.IntValue
				Case SqlDbType.Decimal, SqlDbType.Real, SqlDbType.Float, SqlDbType.Money, SqlDbType.SmallMoney
					Return Field.DataCategoryEnum.DecValue
				Case SqlDbType.DateTime, SqlDbType.Date, SqlDbType.DateTime2, SqlDbType.SmallDateTime
					Return Field.DataCategoryEnum.DateValue
				Case SqlDbType.Bit
					Return Field.DataCategoryEnum.BooleanValue
				Case Else
					Return Field.DataCategoryEnum.OtherValue
			End Select
		Catch ex As Exception
			Return Field.DataCategoryEnum.OtherValue
		End Try
	End Function

End Module
