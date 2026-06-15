# DOit Finances

## Descrição

O **DOit Finances** é um dashboard financeiro interativo para análise de planilhas Excel exportadas do DOit ERP. A aplicação permite visualizar indicadores-chave de performance (KPIs), gráficos financeiros detalhados e tabelas com dados completos — tudo diretamente no navegador, sem necessidade de servidor ou instalação complexa.

O objetivo é fornecer uma ferramenta prática e acessível para gestores e analistas financeiros que utilizam o DOit ERP como sistema de gestão, permitindo análises rápidas e exportação de relatórios em PDF.

## Screenshots

![Dashboard DOit Finances](screenshot-placeholder.png)

## Funcionalidades

- **KPI Dashboard** — Indicadores financeiros resumidos com valores totais, médias e comparativos
- **Gráficos financeiros** — Visualizações interativas com Chart.js (barras, linhas, pizza)
- **Tabelas detalhadas** — Dados completos com ordenação e busca
- **Filtros por data e categoria** — Segmentação dos dados por período, categoria, projeto e outros critérios
- **Exportação PDF** — Geração de relatórios em PDF com os dados e gráficos exibidos
- **Análise por projeto** — Visão consolidada por projeto com receitas, despesas e saldo

## Formato de Arquivo Suportado

### Extensões aceitas

- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

### Tamanho máximo

- **50MB** por arquivo

### Colunas esperadas

| Coluna | Obrigatória |
|--------|:-----------:|
| Projeto | Não |
| Cliente | Não |
| Categoria | Não |
| **Valor** | **Sim** |
| **Tipo** | **Sim** |
| **Data** | **Sim** |
| Status | Não |
| Centro de custo | Não |
| Responsável | Não |
| Departamento | Não |
| Conta | Não |

> ⚠️ As colunas **Valor**, **Tipo** e **Data** são obrigatórias. Sem elas, a aplicação não conseguirá processar o arquivo corretamente.

## Instalação

### Método 1: Zero-install (recomendado)

Nenhuma instalação é necessária. Basta abrir o arquivo `index.html` diretamente no navegador:

1. Baixe ou clone o repositório
2. Abra o arquivo `index.html` no seu navegador (duplo-clique ou arraste para o navegador)
3. Pronto! A aplicação funciona imediatamente

### Método 2: Servidor local com npm

Se preferir usar um servidor de desenvolvimento local:

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/doit-finances.git
cd doit-finances

# Instale as dependências de desenvolvimento
npm install

# Inicie o servidor local
npm start
```

O servidor será iniciado em `http://localhost:3000/` e abrirá automaticamente no navegador.

> **Pré-requisito:** [Node.js](https://nodejs.org/) instalado (v14 ou superior).

## Como Usar

A aplicação roda inteiramente no lado do cliente (client-side) — **nenhum servidor é necessário**. Seus dados nunca saem do seu computador.

### Abrindo localmente

1. Abra o arquivo `index.html` diretamente no navegador (via duplo-clique ou pelo menu Arquivo > Abrir)
2. A interface carregará com a área de upload visível

### Importando arquivo Excel

Existem duas formas de importar sua planilha:

- **Drag-and-drop**: Arraste o arquivo Excel diretamente para a área de upload indicada
- **Seletor de arquivo**: Clique na área de upload para abrir o seletor de arquivos do sistema

Após o upload, o dashboard será preenchido automaticamente com os KPIs, gráficos e tabelas baseados nos dados da planilha.

### Usando servidor local (opcional)

Para quem preferir usar um servidor HTTP local:

```bash
npm start
```

Isso iniciará o servidor em `http://localhost:3000/` com abertura automática no navegador.

## Stack Tecnológica

| Tecnologia | Uso |
|-----------|-----|
| **HTML5** | Estrutura da aplicação (single-page) |
| **CSS3** | Estilização com glassmorphism e design responsivo |
| **JavaScript (ES6+)** | Lógica da aplicação, processamento de dados |
| **Chart.js** | Geração de gráficos interativos |
| **SheetJS (xlsx)** | Leitura e parsing de arquivos Excel |
| **jsPDF** | Exportação de relatórios em PDF |

Todas as bibliotecas de terceiros são carregadas via CDN — nenhuma etapa de build é necessária.

## Contribuindo

Contribuições são bem-vindas! Siga os passos abaixo:

### 1. Fork e Clone

```bash
# Faça um fork do repositório no GitHub
# Em seguida, clone o seu fork:
git clone https://github.com/seu-usuario/doit-finances.git
cd doit-finances
```

### 2. Crie uma branch

Use o padrão de nomenclatura `feature/nome-da-feature` para novas funcionalidades:

```bash
git checkout -b feature/nome-da-feature
```

Para correções de bugs, use `fix/descricao-do-bug`:

```bash
git checkout -b fix/descricao-do-bug
```

### 3. Faça suas alterações

- Mantenha o código limpo e bem documentado
- Teste suas alterações abrindo o `index.html` no navegador

### 4. Envie um Pull Request

```bash
git add .
git commit -m "feat: descrição concisa da alteração"
git push origin feature/nome-da-feature
```

Em seguida, abra um **Pull Request** no GitHub descrevendo suas alterações.

## Licença

Este projeto está licenciado sob a **Licença MIT** — veja o arquivo [LICENSE](LICENSE) para detalhes.
