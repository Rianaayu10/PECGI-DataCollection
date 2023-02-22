Imports System.Data.SqlClient
Imports System.Transactions
Public Class clsSchedulerCSV_BA_DB

    Public Shared Function InsertUpdate_BA(ByVal pData As List(Of clsSchedulerCSV_BA)) As Integer

        Dim iResult As Integer = 0
        Dim con As New SqlConnection
        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            For Each pEachData In pData

                cmd = New SqlCommand("sp_BA_Insert_CSV", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = con
                cmd.Transaction = SQLTrans

                cmd.Parameters.AddWithValue("MachineCode", pEachData.MachineCode)
                cmd.Parameters.AddWithValue("ModeCls", pEachData.ModeCls)
                cmd.Parameters.AddWithValue("StatusCls", pEachData.StatusCls)
                cmd.Parameters.AddWithValue("AlarmCode", pEachData.AlarmCode)
                cmd.Parameters.AddWithValue("StartTime", Format(CDate(pEachData.StartTime), "yyyy-MM-dd HH:mm:ss"))
                cmd.Parameters.AddWithValue("EndTime", Format(CDate(pEachData.EndTime), "yyyy-MM-dd HH:mm:ss"))

                iResult = cmd.ExecuteNonQuery()

            Next

            SQLTrans.Commit()

        Catch ex As Exception
            SQLTrans.Rollback()
            iResult = 0
        End Try

        Return iResult

    End Function

    Public Shared Function InsertData_History(ByVal pData As DataTable, ByVal pSPName As String) As Integer

        Dim iResult As Integer = 0
        Dim con As New SqlConnection
        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            For i As Integer = 0 To pData.Rows.Count - 1

                cmd = New SqlCommand("sp_PLC_CSV_tmp", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = con
                cmd.Transaction = SQLTrans

                cmd.Parameters.AddWithValue("MixDate", Format(CDate(pData.Rows(i).Item(0)), "yyyy-MM-dd HH:mm:ss"))
                cmd.Parameters.AddWithValue("Mode_Cls", Trim(pData.Rows(i).Item(2)))
                cmd.Parameters.AddWithValue("Status_Cls", Trim(pData.Rows(i).Item(3)))
                cmd.Parameters.AddWithValue("Alarm_Code_E", Trim(pData.Rows(i).Item(4)))
                cmd.Parameters.AddWithValue("Alarm_Code_F", Trim(pData.Rows(i).Item(5)))
                cmd.Parameters.AddWithValue("Alarm_Code_G", Trim(pData.Rows(i).Item(6)))
                cmd.Parameters.AddWithValue("Alarm_Code_H", Trim(pData.Rows(i).Item(7)))
                cmd.Parameters.AddWithValue("Alarm_Code_I", Trim(pData.Rows(i).Item(8)))
                cmd.Parameters.AddWithValue("Alarm_Code_J", Trim(pData.Rows(i).Item(9)))
                cmd.Parameters.AddWithValue("Alarm_Code_K", Trim(pData.Rows(i).Item(10)))
                cmd.Parameters.AddWithValue("Alarm_Code_L", Trim(pData.Rows(i).Item(11)))
                cmd.Parameters.AddWithValue("Alarm_Code_M", Trim(pData.Rows(i).Item(12)))
                cmd.Parameters.AddWithValue("Alarm_Code_N", Trim(pData.Rows(i).Item(13)))
                cmd.Parameters.AddWithValue("Alarm_Code_O", Trim(pData.Rows(i).Item(14)))
                cmd.Parameters.AddWithValue("Alarm_Code_P", Trim(pData.Rows(i).Item(15)))
                cmd.Parameters.AddWithValue("MchName", pSPName)

                iResult = cmd.ExecuteNonQuery()

            Next

            SQLTrans.Commit()

        Catch ex As Exception
            SQLTrans.Rollback()
            iResult = 0
        End Try

        Return iResult

    End Function

    Public Shared Function InsertData_Info(ByVal pData As DataTable, ByVal pSPName As String) As Integer

        Dim iResult As Integer = 0
        Dim con As New SqlConnection
        con = New SqlConnection(ConStr)
        con.Open()

        Dim cmd As SqlCommand
        Dim SQLTrans As SqlTransaction

        SQLTrans = con.BeginTransaction

        Try

            For i As Integer = 0 To pData.Rows.Count - 1

                cmd = New SqlCommand("sp_PLC_CSV_tmp2", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Connection = con
                cmd.Transaction = SQLTrans

                cmd.Parameters.AddWithValue("MixDate", Format(CDate(pData.Rows(i).Item(0)), "yyyy-MM-dd HH:mm:ss"))
                cmd.Parameters.AddWithValue("Mode_Cls", Trim(pData.Rows(i).Item(2)))
                cmd.Parameters.AddWithValue("Status_Cls", Trim(pData.Rows(i).Item(3)))
                cmd.Parameters.AddWithValue("Alarm_Code_E", Trim(pData.Rows(i).Item(4)))
                cmd.Parameters.AddWithValue("Alarm_Code_F", Trim(pData.Rows(i).Item(5)))
                cmd.Parameters.AddWithValue("Alarm_Code_G", Trim(pData.Rows(i).Item(6)))
                cmd.Parameters.AddWithValue("Alarm_Code_H", Trim(pData.Rows(i).Item(7)))
                cmd.Parameters.AddWithValue("Alarm_Code_I", Trim(pData.Rows(i).Item(8)))
                cmd.Parameters.AddWithValue("Alarm_Code_J", Trim(pData.Rows(i).Item(9)))
                cmd.Parameters.AddWithValue("Alarm_Code_K", Trim(pData.Rows(i).Item(10)))
                cmd.Parameters.AddWithValue("Alarm_Code_L", Trim(pData.Rows(i).Item(11)))
                cmd.Parameters.AddWithValue("Alarm_Code_M", Trim(pData.Rows(i).Item(12)))
                cmd.Parameters.AddWithValue("Alarm_Code_N", Trim(pData.Rows(i).Item(13)))
                cmd.Parameters.AddWithValue("Alarm_Code_O", Trim(pData.Rows(i).Item(14)))
                cmd.Parameters.AddWithValue("Alarm_Code_P", Trim(pData.Rows(i).Item(15)))
                cmd.Parameters.AddWithValue("Alarm_Code_Q", Trim(pData.Rows(i).Item(16)))
                cmd.Parameters.AddWithValue("Alarm_Code_R", Trim(pData.Rows(i).Item(17)))
                cmd.Parameters.AddWithValue("Alarm_Code_S", Trim(pData.Rows(i).Item(18)))
                cmd.Parameters.AddWithValue("Alarm_Code_T", Trim(pData.Rows(i).Item(19)))
                cmd.Parameters.AddWithValue("Alarm_Code_U", Trim(pData.Rows(i).Item(20)))
                cmd.Parameters.AddWithValue("MchName", pSPName)

                iResult = cmd.ExecuteNonQuery()

            Next

            SQLTrans.Commit()

        Catch ex As Exception
            SQLTrans.Rollback()
            iResult = 0
        End Try

        Return iResult

    End Function

End Class
