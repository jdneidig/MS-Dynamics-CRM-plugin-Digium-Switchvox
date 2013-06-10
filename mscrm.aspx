<%@ Page Language="VB" Debug="true" Explicit="true"%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.WebControls.WebControl" %>

<script runat="server">


Private Sub Page_Load(ByVal sender As System.Object, ByVal E As System.EventArgs)

        Dim strQuerystring As String
        strQuerystring = Request.QueryString("variable")
	Dim strWildCard As String = "%"
	strQuerystring = strWildCard & strQuerystring & strWildCard
	
        
        Dim connectionString As String
        connectionString = "server=sq11;uid=sa;pwd=password;database=CRM"
 
        
        Dim mySqlConnection As SqlConnection
        mySqlConnection = New SqlConnection(connectionString)
 
        
        Dim selectString As String
        selectString = "SELECT ContactId from ContactExtensionBase Where Srobo_BusinessPhoneVOIP LIKE '" & strQuerystring & "' or Srobo_MobilePhoneVOIP LIKE '" & strQuerystring & "'" 
	 
 
       
        Dim mySqlCommand As SqlCommand
        mySqlCommand = mySqlConnection.CreateCommand()
 
        
        mySqlCommand.CommandText = selectString
 
               


mySqlConnection.Open()
	
    
Dim sqlComm As SqlCommand
sqlComm = New SqlCommand(selectString,mySqlConnection)


Dim myDataString As String = Convert.ToString(sqlComm.ExecuteScalar())

    
       
Dim sb As New StringBuilder


sb.Append("<script> var NewWindow;")
sb.Append("NewWindow = window.open(")
sb.Append("'http://sq11/sfa/conts/edit.aspx?id=" & myDataString)
sb.Append("" & "','Report','toolbar=no,width=800,height=675,menubar=no,resizable=yes,maximize=yes'); NewWindow.focus();")
sb.Append(chr(60) & "/script>")

Response.Write(sb.ToString())

mySqlConnection.Close()

End Sub

</script>


<img src="..\_imgs\logo_dynamics_crm_featured.gif" width="170" height="135" />






	
    