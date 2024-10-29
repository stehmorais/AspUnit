# AspUnit
Apresentação que realizei sobre a ferramenta de testes unitários da matéria de Qualidade e Testes de Software - FATEC<br><br>
![image](https://github.com/user-attachments/assets/382c6692-6695-4ab9-8af0-12522c5be402)


## Projeto de Validação de CNPJ em ASP
Este projeto consiste em um sistema de validação de CNPJ desenvolvido em ASP (Active Server Pages), com uma série de testes automatizados para verificar a corretude do código de validação. A aplicação utiliza uma função chamada ValidateCNPJ, que verifica se o CNPJ fornecido segue o formato e regras de validação brasileiras, incluindo os dígitos verificadores.

## Funcionalidades Principais<br>

<b>Função de Validação de CNPJ (ValidateCNPJ):
Esta função realiza uma série de verificações:</b><br>

<b>Valida o formato do CNPJ</b> (espera que esteja no padrão ##.###.###/####-##).<br>
<b>Remove pontuações e espaços para garantir que</b> o CNPJ tenha exatamente 14 dígitos numéricos.<br>
<b>Calcula os dígitos verificadores</b> para confirmar a validade do CNPJ.</b><br>

## Testes Automatizados:</b><br>
O projeto inclui um conjunto de testes para validar a função ValidateCNPJ e demonstrar o uso de diferentes tipos de asserções:</b>

<b>Testes de CNPJ Inválido:</b> Verifica o comportamento da função com CNPJs inválidos, incluindo CNPJs com caracteres incorretos e formatos incorretos.<br>
<b>Outros Testes de Asserção:</b> Exemplos de testes genéricos foram incluídos para mostrar o uso das asserções AssertEquals, AssertNotEquals, AssertTrue e AssertExists, ajudando a garantir a qualidade do código.

## Estrutura do Projeto
<b>cnpj.asp:</b> Contém a função principal ValidateCNPJ, que implementa a lógica de validação.<br>
<b>Testes de Unidade:</b> As funções de teste estão implementadas para validar diferentes cenários. Cada teste verifica um aspecto específico da funcionalidade de validação de CNPJ, permitindo identificar rapidamente problemas no código.

## Tecnologias e Dependências
<b>ASP Clássico:</b> Usado para a lógica de backend e execução dos testes.<br>
<b>VBScript:</b> Linguagem utilizada no ASP para criar funções e rotinas de validação e teste.<br>
<b>Serviço IIS:</b> Para rodar o projeto localmente, é necessário um servidor IIS configurado, com suporte para ASP.

### Como Executar

- Instale o IIS Express, somente para windows

- Abra o CMD

- Vá para sua instalação: cd \Program Files\IIS Express

- Execute o comando: iisexpress /path: "caminho para a pasta com o código"

- Acesse o localhost na porta 8080 (localhost:8080)

- Execute o arquivo de testes

- Clique em Run Tests na página aberta
