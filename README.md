# Automação de Pacientes — Extração de Exames por ID e Dia

## Propósito

1. Ler todas as planilhas em `confidenciais/planilha/` e todos os PDFs de exames em `confidenciais/pdfs/` (recursivo).
2. Identificar a data de cada planilha pela linha `Data: DD/MM/AAAA` e a data de cada página do exame por `Data do exame`.
3. Comparar o **nome** de cada paciente da planilha com os nomes encontrados nas páginas dos PDFs **do mesmo dia**
   (ignorando acentos, pontuação e variações de digitação).
4. Para cada paciente com correspondência de **alta confiança**, extrair a página correspondente e salvar como
   `confidenciais/separados/dia-DD-MM/[id].pdf`.
5. Não sobrescrever arquivos já existentes: a regra de "começar do ID 7" era do primeiro lote (dia 18);
   agora cada dia gera apenas os IDs que ainda faltam na pasta daquele dia.
6. Gerar um relatório `confidenciais/relatorio_nao_encontrados.txt` com os pacientes sem correspondência confiável.

## Como usar

1. Coloque os PDFs de exames em `confidenciais/pdfs/`, organizados por mês se quiser
   (ex.: `confidenciais/pdfs/setembro/Dia 17.pdf`, `Parte 01 - Dia 18.pdf`...).
2. Coloque uma planilha por dia em `confidenciais/planilha/`, com a linha `Data: DD/MM/AAAA`
   e as colunas `Nº` e `Nome` (a data define a pasta de saída).
3. Execute:

   ```bash
   python script.py
   ```

4. Verifique a saída:
   - Arquivos `[id].pdf` gerados em `confidenciais/separados/dia-DD-MM/`.
   - `confidenciais/relatorio_nao_encontrados.txt` com pacientes que precisam de revisão manual.

## Critérios de correspondência

- Nome **normalizado idêntico** (sem acentos/pontuação) — aceito sempre.
- OU semelhança (`difflib`) >= 0.90 com folga de >= 0.08 sobre o 2º candidato — aceito como alta confiança.
- Caso contrário, o paciente é listado no relatório para conferência manual.

## Estrutura

```
script.py            # automação
README.md
.gitignore           # ignora confidenciais/ e __pycache__
confidenciais/       # dados sensíveis (fora do git)
  pdfs/              # PDFs-fonte dos exames, por mês
    setembro/
      Dia 17.pdf
      Parte 01 - Dia 18.pdf
      Parte 02 - Dia 18.pdf
  planilha/          # planilha de cada dia (Data + Nº/Nome)
  separados/         # saída, separada por dia
    dia-18-09/
    dia-17-09/
  relatorio_nao_encontrados.txt   # gerado automaticamente
```