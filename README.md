# Automação de Relatórios Patrimoniais

Este repositório traz um conjunto de scripts e um modelo de dados pensados
para acelerar a produção de relatórios técnicos de inspeção, tabelas de
fotos e planilhas de acompanhamento em Word e Excel. A proposta é servir
como um **template inteligente** que você pode adaptar aos diferentes
ativos do seu portfólio, mantendo uma apresentação padronizada e fácil de
atualizar.

## Visão geral da solução

- **`automation/report_generator.py`**: script principal que lê um arquivo
  JSON com os dados do ativo, gera um relatório técnico em Word e uma
  planilha de inventário em Excel. O relatório inclui tabelas formatadas,
  seções de inspeção, resumo executivo e registro fotográfico com imagens
  anotadas (setas, caixas, textos).
- **`data/sample_asset.json`**: exemplo de estrutura de dados esperada. Use
  este arquivo como base para montar os levantamentos reais.
- **Pasta `media/`** (crie conforme necessidade): onde devem ficar as fotos
  originais utilizadas nas inspeções.
- **Pasta `output/`**: destino padrão dos arquivos gerados. As imagens
  anotadas são armazenadas dentro de `output/annotated/`.

## Requisitos

Instale as dependências em um ambiente virtual com Python 3.9+:

```bash
pip install python-docx openpyxl pillow
```

## Estrutura do arquivo JSON

O arquivo de dados combina informações do ativo, resumo executivo, fotos e
blocos de inspeção. Campos principais:

```json
{
  "asset": {
    "name": "Nome do ativo",
    "code": "Código interno",
    "address": "Endereço completo",
    "inspection_date": "AAAA-MM-DD",
    "responsible": "Responsável técnico"
  },
  "summary": [
    {"topic": "Estrutura", "notes": "Principais achados"}
  ],
  "photos": [
    {
      "file": "fachada.jpg",
      "caption": "Legenda exibida no relatório",
      "annotations": [
        {"type": "arrow", "from": [120, 480], "to": [260, 340]},
        {"type": "box", "top_left": [300, 250], "bottom_right": [420, 360]},
        {"type": "text", "label": "Observação", "position": [310, 220]}
      ]
    }
  ],
  "inspections": [
    {
      "location": "Cobertura",
      "elements": [
        {
          "element": "Manta impermeabilizante",
          "status": "Necessita manutenção",
          "priority": "Alta",
          "recommendation": "Ação recomendada",
          "photo": "cobertura.jpg"
        }
      ]
    }
  ]
}
```

## Como usar

1. Ajuste o JSON para refletir os dados coletados no levantamento.
2. Salve as imagens em uma pasta (por padrão `media/`).
3. Execute o gerador apontando para o arquivo de dados:

   ```bash
   python automation/report_generator.py data/sample_asset.json \
       --media-dir media \
       --output-dir output
   ```

4. Os arquivos gerados serão:
   - `output/relatorio_tecnico.docx`
   - `output/inventario_inspecao.xlsx`
   - imagens anotadas dentro de `output/annotated/`

### Personalizações rápidas

- Use a flag `--sem-anotacoes` caso queira inserir as fotos originais sem
  sobreposições.
- Amplie o script adicionando campos extras nas tabelas ou alterando estilos
  (por exemplo, substitua o tema da tabela por outro estilo do Word).
- Centralize a base de dados em uma planilha ou sistema e gere o JSON com
  scripts auxiliares, mantendo sempre a mesma estrutura.

## Dicas para fluxos em equipe

- Guarde o JSON e as fotos originais junto com o relatório final. Assim,
  qualquer revisão futura pode ser refeita automaticamente.
- Versione o arquivo de dados para acompanhar a evolução das inspeções.
- Utilize o Excel gerado como checklist dinâmico, filtrando por prioridade e
  alimentando dashboards de manutenção.

Com esse setup você transforma horas de ajustes manuais em uma operação
padronizada, replicável e pronta para ser incrementada com novos recursos
(geração de PDF, dashboards automatizados etc.).
