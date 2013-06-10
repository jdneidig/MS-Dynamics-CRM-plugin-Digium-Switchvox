MS-Dynamics-CRM-plugin-Digium-Switchvox
=======================================

Microsoft Dynamics CRM plugin for Digium Switchvox

source code for connecting and querying incoming calls from Switchvox to the MS Dynamics CRM database. It is very simplistic and if you are familiar with Switchvox, ASP.net, and MS Dynamics CRM you’ll find it very easy to implement. I have attached the source code, it is an asp.net page. There are some customized fields you will have to add to your CRM database, or you can use the existing phone numbers under contacts. You’ll see in the code we are querying against customized fields. Basically, you will copy and past the mscrm.aspx file in the root folder of where your website is for CRM on your network. This has been tested for MS Dynamics CRM 4.0. We have not tested it on earlier versions but it should work. And you will see in one of the screen shots below you will point to where you copied the mscrm.aspx file to on your CRM server "http://[your server name:port]/mscrm.aspx" and then you’ll add the “?variable=%CALLER_ID_NUMBER%” so that when an incoming call comes through your switchvox client it will query the database for that number and bring up the contacts information.
