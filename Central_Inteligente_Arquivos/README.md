# Central Inteligente de Arquivos

Solução local para Windows com Excel Desktop, VBA, PowerShell, Word e Microsoft 365 para inventariar, diagnosticar, pesquisar e exportar informações de arquivos em pastas locais ou OneDrive.

## Princípios de segurança

- Não exclui arquivos originais.
- Não move arquivos originais.
- Não renomeia arquivos originais.
- Não altera o conteúdo dos arquivos analisados.
- Gera saídas somente na pasta configurada.
- Registra logs técnicos e operacionais.
- Solicita confirmação explícita para desbloqueio de arquivos confiáveis.

## Estrutura

- `01_Excel`: modelo das abas e instruções para criar o arquivo `.xlsm`.
- `02_VBA`: módulos VBA importáveis no Excel Desktop.
- `03_PowerShell`: scripts de inventário, diagnóstico, extração, desbloqueio e relatórios.
- `04_Saidas`: pastas para CSV, HTML, Word, Excel e logs.
- `05_Documentacao`: manuais, checklist, erros comuns e plano de testes.

## Fluxo resumido

1. Crie o arquivo `Central_Inteligente_Arquivos.xlsm` no Excel Desktop.
2. Importe os módulos `.bas` da pasta `02_VBA`.
3. Execute `InicializarCentral` para criar abas, cabeçalhos e botões.
4. No painel, escolha uma pasta local ou OneDrive.
5. Execute inventário, diagnóstico, extração, busca e exportações.

## Requisitos

- Windows 10/11.
- Excel Desktop com macros habilitadas para local confiável.
- PowerShell 5.1 ou superior.
- Word Desktop para exportação `.docx` via automação COM.
- Microsoft 365 instalado para extração de Word, Excel e PowerPoint via COM.
