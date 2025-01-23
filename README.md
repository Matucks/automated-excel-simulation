# Automacão de Planilhas Excel

Este repositório contém um projeto em Python para a automação de manipulação e análise de dados em planilhas Excel. O objetivo principal é processar dados de entrada, realizar cálculos e gerar uma planilha formatada e consolidada com resultados detalhados.

---

## Funcionalidades Principais

### 1. **Processamento Automático**
- Importa arquivos Excel de entrada e metas de uma pasta específica.
- Remove linhas desnecessárias, como "Total Consolidado".
- Realiza cálculos, incluindo:
  - Soma de colunas.
  - Diferença entre metas e totais.
  - Percentual de alcance.

### 2. **Consolidação de Dados**
- Gera uma aba consolidada com todas as informações, incluindo totais e percentuais.
- Adiciona uma linha de "Total Consolidado" com somatórios e cálculos gerais.

### 3. **Formatação e Estilo**
- Aplica estilos personalizados aos cabeçalhos e células.
- Insere barras de dados na coluna de percentual de alcance (% ALCANCE).
- Ajusta o layout para melhor visualização dos resultados.

---

## Estrutura do Projeto

### 1. **Caminhos de Entrada e Saída**
- **entrada_path**: Pasta onde são armazenados os arquivos de dados.
- **objetivo_path**: Pasta onde estão armazenadas as metas.
- **output_path**: Local onde o arquivo consolidado será salvo.

### 2. **Funções Principais**
- **encontrar_arquivo_base**: Busca o primeiro arquivo Excel em uma pasta especificada.
- **carregar_dados**: Carrega e filtra os dados de entrada e metas.
- **criar_e_preencher_planilha**: Realiza o processamento principal e gera o DataFrame consolidado.
- **salvar_planilha_com_estilo**: Aplica formatação e salva o arquivo final.

---

## Configuração

### 1. **Requisitos**
- **Python 3.8 ou superior.**

### 2. **Bibliotecas Necessárias**
- `pandas`
- `openpyxl`

Instale as dependências executando o comando abaixo:
```bash
pip install pandas openpyxl
```

### 3. **Configuração dos Caminhos**
Ajuste os valores de `entrada_path`, `objetivo_path` e `output_path` no código para refletir os diretórios da sua máquina.

---

## Como Usar

1. **Prepare os Arquivos**
   - Coloque os arquivos de dados e metas nas pastas configuradas (`entrada_path` e `objetivo_path`).

2. **Execute o Script**
   ```bash
   python nome_do_script.py
   ```

3. **Verifique os Resultados**
   - O arquivo consolidado será salvo no diretório configurado no `output_path`.

---

## Estrutura da Planilha Final

### **Colunas Geradas**
- **SETOR**: Nome do setor analisado.
- **META**: Meta estabelecida para cada setor.
- **CONCLUÍDO**: Total já realizado.
- **EM ANDAMENTO**: Itens pendentes de conclusão.
- **APROV. FINANCEIRA**: Itens aprovados financeiramente.
- **APROV. GERAL**: Itens aprovados em geral.

---

## Contribuições

Contribuições são bem-vindas! Para relatar problemas, sugerir melhorias ou enviar pull requests, utilize a aba ["Issues"](https://github.com/seu-usuario/excel-automation-project/issues) no repositório.

---

## Licença

Este projeto está licenciado sob a [MIT License](https://opensource.org/licenses/MIT).

---

## Autor

- **Gabriel Matuck**  
  - **E-mail**: [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

