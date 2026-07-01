# Plano de Testes

## Massa mínima

- TXT com texto simples.
- CSV com acentos.
- DOCX com parágrafos.
- XLSX com duas abas.
- PPTX com caixas de texto.
- PDF pesquisável.
- PDF digitalizado.
- Imagem.
- Arquivo 0 KB.
- Arquivo sem extensão.
- Arquivos duplicados.
- Arquivo temporário `~$`.
- Arquivo com caminho longo.

## Casos

1. Inicializar abas e botões.
2. Escolher pasta.
3. Executar inventário com subpastas.
4. Executar inventário sem subpastas.
5. Ignorar ocultos.
6. Ignorar temporários.
7. Detectar duplicados por nome, tamanho e hash.
8. Detectar alertas.
9. Extrair conteúdo compatível.
10. Tratar formatos incompatíveis com mensagem amigável.
11. Pesquisar palavra-chave.
12. Exportar Word.
13. Exportar Excel.
14. Gerar HTML.
15. Listar bloqueados sem desbloquear.
16. Desbloquear somente após confirmação explícita.
