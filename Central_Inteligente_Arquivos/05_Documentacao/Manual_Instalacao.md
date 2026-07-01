# Manual de Instalação

## Pré-requisitos

- Windows 10 ou 11.
- Excel Desktop com suporte a VBA.
- Word Desktop para exportação em `.docx`.
- PowerPoint Desktop para extração de apresentações.
- PowerShell 5.1 ou superior.

## Instalação

1. Copie a pasta `Central_Inteligente_Arquivos` para um local de trabalho confiável.
2. Abra o Excel e crie `01_Excel/Central_Inteligente_Arquivos.xlsm`.
3. Importe os módulos `.bas` da pasta `02_VBA`.
4. Execute `InicializarCentral`.
5. Configure a pasta como local confiável no Excel, se necessário.

## Política de execução

Os scripts são chamados pelo VBA com `-ExecutionPolicy Bypass` somente no processo iniciado. Isso não altera a política global do computador.
