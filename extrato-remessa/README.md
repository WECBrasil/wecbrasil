# Extrato de Remessa

Este projeto é uma aplicação web para processamento de extratos de remessa. Ele permite o upload de planilhas, visualização de dados e exportação de resultados em formato PDF.

## Estrutura do Projeto

- **index.html**: Documento HTML principal da aplicação. Inclui links para folhas de estilo e scripts, e configura a estrutura para upload de arquivos, visualização de dados e exportação de resultados.
  
- **app.js**: Contém a classe `ExtratoProcessor`, que gerencia a lógica para upload de arquivos, processamento de dados, exibição de pré-visualizações e exportação de resultados como PDFs. Inclui métodos para gerenciar valores negativos, gerar HTML para PDFs e formatar dados.

- **style.css**: Contém os estilos para a aplicação, incluindo layout, designs de cartões, tabelas, botões e elementos de design responsivo.

- **vendor/**: Diretório que contém bibliotecas de terceiros:
  - **bootstrap.min.css**: CSS minificado do Bootstrap, fornecendo estilos e componentes para design responsivo.
  - **bootstrap.bundle.min.js**: Pacote JavaScript minificado do Bootstrap, que inclui os scripts necessários para componentes do Bootstrap.
  - **xlsx.min.js**: Biblioteca JavaScript minificada para manipulação de arquivos Excel, permitindo que a aplicação leia e processe dados de planilhas.
  - **html2pdf.bundle.min.js**: Biblioteca JavaScript minificada para conversão de conteúdo HTML em formato PDF.
  - **jszip.min.js**: Biblioteca JavaScript minificada para criação e gerenciamento de arquivos ZIP, usada para exportar múltiplos PDFs.

- **assets/logos/**: Diretório destinado ao armazenamento de imagens de logotipos que podem ser carregadas para uso na aplicação.

- **.vscode/launch.json**: Contém configurações de depuração para a aplicação em um ambiente de desenvolvimento.

- **package.json**: Arquivo de configuração para npm, listando as dependências do projeto, scripts e metadados.

## Instruções de Uso

1. **Instalação**: Clone o repositório e abra o projeto em um servidor local.
2. **Upload de Arquivos**: Utilize a seção de upload para enviar planilhas no formato `.xlsx`, `.xls` ou `.csv`.
3. **Visualização**: Após o upload, visualize os dados e gerencie valores negativos, se necessário.
4. **Exportação**: Configure a exportação e gere arquivos PDF agrupados por departamento.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.