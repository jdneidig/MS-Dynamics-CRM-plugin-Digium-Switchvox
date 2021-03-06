<%@ Page Language="VB" Explicit="true"%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.WebControls.WebControl" %>

<script runat="server">


Private Sub Page_Load(ByVal sender As System.Object, ByVal E As System.EventArgs)

        Dim strQuerystring As String
        strQuerystring = Request.QueryString("variable")
        
        Dim myIDString As String = New String("")
        Dim myNameString As String = New String("")

        Dim myPhoneNumber As String = strQuerystring

        Dim strWildCard As String = "%"

        strQuerystring = strWildCard & Left(strQuerystring, 3) & strWildCard & Mid(strQuerystring, 4, 3) & strWildCard & Mid(strQuerystring, 7, 4) & strWildCard


        Dim connectionString As String

        connectionString = "server=SQL1;uid=SV2CRM_User;pwd=mypassword;database=myCRM_MSCRM;"

        Dim mySqlConnection As SqlConnection

        mySqlConnection = New SqlConnection(connectionString)

        Dim selectAccountID As String
        Dim selectAccountName As String

        Dim selectContactID As String
        Dim selectContactName As String


        selectAccountID = "select AccountId from AccountBase where Telephone1 like '" & strQuerystring & "' or Telephone2 like '" & strQuerystring & "' or Telephone3 like '" & strQuerystring & "'"




        'Did these do anything?
        'Dim mySqlCommand As SqlCommand

        'mySqlCommand = mySqlConnection.CreateCommand()

        'mySqlCommand.CommandText = selectAccount

        mySqlConnection.Open()

        Dim sqlComm As SqlCommand

        sqlComm = New SqlCommand(selectAccountID, mySqlConnection)

        myIDString = Convert.ToString(sqlComm.ExecuteScalar)


        Dim myAccountOrContactString As String = "accts"
        Dim myAccountOrContactInteger As Integer = 1


        'Dim myAccount As String = Convert.ToString(sqlComm.ExecuteScalar)

        If myIDString.Length = 0 Then

            'No Account found then try to find a Contact
            selectContactID = "select ContactId  from ContactBase where Telephone1 like '" & strQuerystring & "' or Telephone2 like '" & strQuerystring & "' or Telephone3 like '" & strQuerystring & "' or MobilePhone like '" & strQuerystring & "'"

            sqlComm.CommandText = selectContactID
            myIDString = Convert.ToString(sqlComm.ExecuteScalar)
            If myIDString.Length > 0 Then

                selectContactName = "select FullName from ContactBase where ContactId  = '" & myIDString & "'"
                sqlComm.CommandText = selectContactName
                myAccountOrContactString = "conts"
                myAccountOrContactInteger = 2

            End If




        Else
            selectAccountName = "select Name from AccountBase where AccountId = '" & myIDString & "'"
            sqlComm.CommandText = selectAccountName
            'myNameString = Convert.ToString(sqlComm.ExecuteScalar)
            myAccountOrContactString = "accts"
            myAccountOrContactInteger = 1

        End If

        If myIDString.Length > 0 Then

            myNameString = Convert.ToString(sqlComm.ExecuteScalar)


            If myNameString.Length > 0 Then

                'With myNameString
                myNameString = myNameString.Replace(" ", "%20")
                myNameString = myNameString.Replace("&", "%26")
                myNameString = myNameString.Replace(",", "%2c")
                'End With



            End If

            myPhoneNumber = "%28" & Left(myPhoneNumber, 3) & "%29%20" & Mid(myPhoneNumber, 4, 3) & "-" & Mid(myPhoneNumber, 7, 4)

            'With myPhoneNumber
            '    .Replace("(", "%28")
            '    .Replace(")", "%29")
            '    .Replace(" ", "%20")
            'End With




            Dim sbPhoneCall As New StringBuilder

            sbPhoneCall.Append("<script> var NewWindow;")
            sbPhoneCall.Append("NewWindow = window.open(")
            sbPhoneCall.Append("'http://sql1:5555/myCRM/activities/phone/edit.aspx?pId={" & myIDString & "}")
            sbPhoneCall.Append("&pType=" & myAccountOrContactInteger & "&pName=" & myNameString)
            sbPhoneCall.Append("&partyid={" & myIDString & "}")
            sbPhoneCall.Append("&partytype=" & myAccountOrContactInteger & "&partyname=" & myNameString)
            sbPhoneCall.Append("&partyaddressused=&contactInfo=" & myPhoneNumber)
            sbPhoneCall.Append("" & "','Report','toolbar=no,width=800,height=675,menubar=no,resizable=yes,maximize=yes'); NewWindow.focus();")
            sbPhoneCall.Append(Chr(60) & "/script>")

            Response.Write(sbPhoneCall.ToString())

            'Dim myTest As String = sbAccountOrContact.ToString()
            'TextBox2.Text = sbAccountOrContact.ToString()


        Else


        End If

        mySqlConnection.Close()


End Sub

</script>

<img src="..\_imgs\dynamics_logo.jpg" width="240" height="240" />






	
    