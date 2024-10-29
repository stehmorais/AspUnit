
<!-- #include file="cnpj.asp"-->
<%

Class CNPJValidatorTest
    Public Function TestCaseNames()
        TestCaseNames = Array("cnpjValido", "cnpjInvalido", "cnpjComCaracteresInvalidos", "cnpjFormatoInvalido", "testeFalho", "testeAssertEquals", "testeAssertNotEquals", "testeAssertTrue", "testeAssertExists")
    End Function

    Public Sub SetUp()
    End Sub

    Public Sub TearDown()
    End Sub

    ' Teste para CNPJ válido
    Public Sub cnpjValido(oTestResult)
        Dim cnpj, resultado
        cnpj = "12.345.678/0001-95"  ' Exemplo de CNPJ válido
        resultado = ValidateCNPJ(cnpj)
        oTestResult.AssertEquals "CNPJ valido", resultado, "Deve ser um CNPJ valido / "
    End Sub

    ' Teste para CNPJ inválido
    Public Sub cnpjInvalido(oTestResult)
        Dim cnpj, resultado
        cnpj = "12.345.678/0001-00"  ' Exemplo de CNPJ inválido
        resultado = ValidateCNPJ(cnpj)
        oTestResult.AssertEquals "CNPJ invalido", resultado, "Deve ser um CNPJ invalido / "
    End Sub

    ' Teste para CNPJ com caracteres inválidos
    Public Sub cnpjComCaracteresInvalidos(oTestResult)
        Dim cnpj, resultado
        cnpj = "12.345.678/0001-9A"  ' Contém caractere não numérico
        resultado = ValidateCNPJ(cnpj)
        oTestResult.AssertEquals "CNPJ invalido", resultado, "Deve ser um CNPJ invalido devido a caracteres invalidos / "
    End Sub

    ' Teste para CNPJ com formato inválido
    Public Sub cnpjFormatoInvalido(oTestResult)
        Dim cnpj, resultado
        cnpj = "12345678000195"  ' Formato incorreto (sem pontuação)
        resultado = ValidateCNPJ(cnpj)
        oTestResult.AssertEquals "CNPJ invalido", resultado, "Deve ser um CNPJ invalido devido ao formato incorreto / "
    End Sub

    Public Sub testeFalho(oTestResult)
        Dim cnpj, resultado
        cnpj = "12.345.678/0001-95"
        resultado = ValidateCNPJ(cnpj)
        oTestResult.AssertEquals "CNPJ invalido", resultado, "Apenas para demonstrar um teste com falha / "
    End Sub

    Public Sub testeAssertEquals(oTestResult)
        Dim expected, actual
        expected = 10
        actual = 10
        oTestResult.AssertEquals expected, actual, "10 = 10, Nao deve falhar!"
    End Sub

    Public Sub testeAssertNotEquals(oTestResult)
        Dim expected, actual
        expected = 10
        actual = 20
        oTestResult.AssertNotEquals expected, actual, "10 != 20, Deve passar!"
    End Sub

    Public Sub testeAssertTrue(oTestResult)
        Dim condition
        condition = (5 * 2 = 10)
        oTestResult.Assert condition, "5 * 2 = 10, Nao deve falhar!"
    End Sub

    Public Sub testeAssertExists(oTestResult)
        Dim obj
        Set obj = Server.CreateObject("Scripting.Dictionary")
        oTestResult.AssertExists obj, "Objeto deve existir."
    End Sub

End Class
%>
