<%
Option Explicit
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<!-- #include file="aspunit/ASPUnitRunner.asp"-->
<!-- #include file="ValidarCNPJ.asp"-->

<%
    Dim Tester
    Set Tester = New UnitRunner
    Tester.AddTestContainer New CNPJValidatorTest
    
    Tester.Display()
%>
