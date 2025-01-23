Automação de Planilhas Excel

Este repositório contém um projeto em Python para a automação de manipulação e análise de dados em planilhas Excel. O objetivo principal é processar dados de entrada, realizar cálculos, e gerar uma planilha formatada e consolidada com resultados detalhados.

Funcionalidades

Processamento Automático:

Importa arquivos Excel de entrada e metas de uma pasta específica.

Remove linhas desnecessárias, como "Total Consolidado".

Realiza cálculos como soma de colunas, diferença entre metas e totais, e percentual de alcance.

Consolidação de Dados:

Gera uma aba consolidada com todas as informações, incluindo totais e percentuais.

Adiciona uma linha de "Total Consolidado" com somatórios e cálculos gerais.

Formatação e Estilo:

Aplica estilos personalizados aos cabeçalhos e células.

Insere barras de dados na coluna de percentual de alcance (% ALCANCE).

Ajusta o layout para melhor visualização dos resultados.

Estrutura do Projeto

Caminhos de Entrada e Saída:

entrada_path: Pasta onde são armazenados os arquivos de dados.

objetivo_path: Pasta onde estão armazenadas as metas.

output_path: Local onde o arquivo consolidado será salvo.

Funções Principais:

encontrar_arquivo_base: Busca o primeiro arquivo Excel em uma pasta especificada.

carregar_dados: Carrega e filtra os dados de entrada e metas.

criar_e_preencher_planilha: Realiza o processamento principal e gera o DataFrame consolidado.

salvar_planilha_com_estilo: Aplica formatação e salva o arquivo final.

Configuração

Requisitos:

Python 3.8 ou superior

Bibliotecas Python:

pandas

openpyxl

Instalação das Dependências:

pip install pandas openpyxl

Configuração dos Caminhos:

Ajuste os valores de entrada_path, objetivo_path e output_path no código para refletir os diretórios da sua máquina.

Como Usar

Coloque os arquivos de dados e metas nas pastas configuradas (entrada_path e objetivo_path).

Execute o script Python:

python nome_do_script.py

Verifique o arquivo gerado no diretório de saída (output_path).

Estrutura da Planilha Final

Colunas Geradas:

SETOR: Nome do setor analisado.

META: Meta estabelecida para cada setor.

CONCLUÍDO: Total já realizado.

EM ANDAMENTO: Itens pendentes de conclusão.

APROV. FINANCEIRA: Itens aprovados financeiramente.

APROV. GERAL: Itens aprovados em geral.

TOTAL: Soma de todas as colunas numéricas anteriores.

DIFERENÇA META: Diferença entre a meta e o total.

% ALCANCE: Percentual do total alcançado em relação à meta.

Linha de Total Consolidado:

Soma de todas as colunas numéricas e cálculos gerais para os resultados consolidados.

Exemplo de Uso

Adicione um arquivo chamado dados.xlsx na pasta de entrada.

Adicione um arquivo chamado metas.xlsx na pasta de metas.

Execute o script e verifique a saída consolidada e formatada.

Contribuição

Contribuições são bem-vindas! Caso encontre problemas ou tenha sugestões de melhorias, sinta-se à vontade para abrir uma issue ou criar um pull request.

Licença

Este projeto está licenciado sob a Licença MIT. 

Consulte o arquivo LICENSE para mais informações.
