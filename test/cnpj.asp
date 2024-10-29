<%
Function ValidateCNPJ(cnpj)
    ' Verifica se o formato do CNPJ é válido usando expressão regular
    Dim regEx, matches
    Set regEx = New RegExp
    regEx.Pattern = "^\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}$" ' Formato: 12.345.678/0001-95
    regEx.IgnoreCase = True
    regEx.Global = False
    
    ' Se o CNPJ não corresponder ao padrão, retorna inválido
    If Not regEx.Test(cnpj) Then
        ValidateCNPJ = "CNPJ invalido"
        Exit Function
    End If

    ' Remove caracteres de pontuação para validação dos dígitos
    cnpj = Replace(Replace(Replace(Replace(cnpj, ".", ""), "/", ""), "-", ""), " ", "")

    ' Verifica se o CNPJ possui 14 dígitos numéricos
    If Len(cnpj) <> 14 Or Not IsNumeric(cnpj) Then
        ValidateCNPJ = "CNPJ invalido"
        Exit Function
    End If

    ' Cálculo do dígito verificador do CNPJ
    Dim pesos1, pesos2, i, soma, digito1, digito2
    pesos1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    pesos2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    ' Calcula o primeiro dígito verificador
    soma = 0
    For i = 1 To 12
        soma = soma + CInt(Mid(cnpj, i, 1)) * pesos1(i - 1)
    Next
    digito1 = (11 - (soma Mod 11))
    If digito1 >= 10 Then digito1 = 0

    ' Calcula o segundo dígito verificador
    soma = 0
    For i = 1 To 13
        soma = soma + CInt(Mid(cnpj, i, 1)) * pesos2(i - 1)
    Next
    digito2 = (11 - (soma Mod 11))
    If digito2 >= 10 Then digito2 = 0

    ' Verifica se os dígitos verificadores correspondem
    If CInt(Mid(cnpj, 13, 1)) = digito1 And CInt(Mid(cnpj, 14, 1)) = digito2 Then
        ValidateCNPJ = "CNPJ valido"
    Else
        ValidateCNPJ = "CNPJ invalido"
    End If

    ' Libera o objeto RegExp
    Set regEx = Nothing
End Function
%>
