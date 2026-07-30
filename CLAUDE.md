> Guia para o Claude Code neste repositório.

## 1. Regra principal: não usar travessões

**Nunca usar `---`, `--` ou `—` (em dash / travessão) em texto corrido.**

Vale para tudo: o conteúdo dos arquivos `.tex`, títulos, tabelas, legendas, mensagens de commit, README e também as respostas no chat.

Em vez de travessão, use a pontuação adequada ao caso:

| Em vez de | Use |
| --- | --- |
| `Gestão 2026 a 2030 --- Plano de Ação` | `Gestão 2026 a 2030: Plano de Ação` |
| `valores ilustrativos --- ajustar depois` | `valores ilustrativos (ajustar depois)` |
| `alinhados à mesma régua --- ambos por ano` | `alinhados à mesma régua, ambos por ano` |
| `Meta Operacional I — Desenvolvimento` | `Meta Operacional I: Desenvolvimento` |

Alternativas aceitas, na ordem de preferência: dois-pontos, parênteses, vírgula, ponto (duas frases) ou reescrever a frase.

Duas exceções, que **não** são texto corrido:

- Separadores de seção em comentários do preâmbulo, que já são a convenção do projeto: `% --- Pacotes básicos`.
- Hífen simples em intervalos e nomes próprios de documento: `2026-2030`, `PE 26-30`.

Se encontrar travessões em texto já existente, avise antes de sair corrigindo em massa.

## 2. Estilo de escrita dos documentos

- Português do Brasil, registro formal institucional (documento oficial da Fundep).
- Não inventar números, prazos ou valores financeiros. Quando o dado não existir, marcar explicitamente como `(a preencher)` ou `(valores ilustrativos)`, como já é feito no plano de ação.
- Preservar a formatação e as cores existentes: os arquivos reproduzem documentos Word originais que estão em [arquivos/](arquivos/).
- Siglas: declarar com `\newacronym` no preâmbulo e usar `\gls{...}` no texto, não a sigla literal.

## 3. Estrutura

Documentos independentes, cada um compilável sozinho:

- [planejamentoEstrategico2030.tex](planejamentoEstrategico2030.tex): relatório de planejamento estratégico de TI. Usa `glossaries-extra` (acrônimos).
- [PE 26-30 - PPA - Plano de Acao - Transformacao Digital - II.tex](PE%2026-30%20-%20PPA%20-%20Plano%20de%20Acao%20-%20Transformacao%20Digital%20-%20II.tex): plano de ação, réplica em LaTeX do `.docx` homônimo, com papel timbrado como fundo de página e paleta de cores fiel ao Word.
- [PE 26-30 - PPA - Plano de Acao - MODELO.tex](PE%2026-30%20-%20PPA%20-%20Plano%20de%20Acao%20-%20MODELO.tex): modelo em branco de plano de ação, baseado no primeiro plano do arquivo acima. Contém identificação, justificativa, atividades, resumo das atividades e cronograma físico-financeiro, com os campos marcados como `(a preencher)`.

Cada plano de ação fica em um arquivo próprio, com uma única AÇÃO. Para criar um novo plano, copiar o modelo e renomear.

Pastas:

- [arquivos/](arquivos/): fontes originais (`.docx`, `.pdf`) que servem de referência. Não editar, é material de origem. Os PDFs daqui são versionados (exceção no `.gitignore`).
- [img/](img/): imagens usadas nos documentos (timbrado, rodapé, organograma).

## 4. Compilação

```bash
latexmk "planejamentoEstrategico2030.tex"
latexmk "PE 26-30 - PPA - Plano de Acao - Transformacao Digital - II.tex"
latexmk -c    # limpa auxiliares
```

- Engine: `pdflatex`, configurada no [.latexmkrc](.latexmkrc).
- O `.latexmkrc` já registra as dependências que chamam `makeglossaries` automaticamente, então não é preciso rodá-lo à mão.
- No macOS, o TeX fica em `/Library/TeX/texbin`. Se `latexmk` não for encontrado, exportar esse caminho no `PATH`.
- No VS Code, a receita compartilhada está em [.vscode/settings.json](.vscode/settings.json) (extensão LaTeX Workshop, build automático ao salvar).

## 5. Versionamento

- Auxiliares LaTeX e PDFs de saída são ignorados pelo `.gitignore`. Não commitar `.aux`, `.log`, `.pdf` da raiz.
- Ao alterar o conteúdo de um documento, atualizar `\ReportVersion` no preâmbulo quando a mudança for significativa.
- Mensagens de commit em português, no imperativo, sem travessões.
