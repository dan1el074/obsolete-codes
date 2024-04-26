# Códigos Obsoletos

Este projeto é uma aplicação em **Node.js** desenvolvida para o setor de engenharia da empresa [Metaro Indústria e Comércio LTDA](https://www.metaro.com.br). Ela realiza uma varredura em um arquivo Excel e extrai códigos que precisam ser desativados no sistema. O arquivo Excel é gerado através de um BI (Business Intelligence) do ERP **Probus**.

## Pré-requisitos

Antes de executar a aplicação, certifique-se de que o seguinte esteja instalado:

- [Node.js](https://nodejs.org/en/download/current) (versão 12 ou superior)
- npm (gerenciador de pacotes do Node.js)

## Instalação

1. Clone este repositório para o seu ambiente local:

    ```bash
    git clone https://github.com/dan1el074/obsolete-codes.git
    ```

2. Navegue até o diretório do projeto:

    ```bash
    cd obsolete-codes
    ```

## Configuração

Coloque o **Arquivo Excel** gerado pelo **Probus** na pasta `obsolete-codes\resource\examples`, depois disso, renomeie o arquivo para `codes.xlsx`

## Uso

Inicie o seguinte arquivo:

```bash
run.bat
```

Ou execute o seguinte comando para iniciar a varredura:

```bash
npm start
