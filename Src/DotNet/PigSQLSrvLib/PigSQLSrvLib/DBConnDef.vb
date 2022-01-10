'**********************************
'* Name: DBConnDef
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Database connection definition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 17/10/2021
'**********************************
Public Class DBConnDef
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub
End Class
