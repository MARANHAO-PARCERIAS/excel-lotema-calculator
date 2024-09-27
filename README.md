
# Calculadora de Repasses LOTEMA

Esta calculadora foi desenvolvida para facilitar o cálculo dos repasses da Loteria do Estado do Maranhão (LOTEMA). Com ela, você pode obter os valores para cada modalidade (Prognóstico Numérico, Quota Fixa e Instantânea) de forma rápida e precisa.

## Funcionalidades

- **Prognóstico Numérico:** Calcula os repasses para diversas modalidades, como Educação, MAPA, Previdência e mais.
- **Quota Fixa:** Calcula os repasses para modalidades específicas, como Educação e Segurança Social.
- **Instantânea:** Permite calcular repasses para modalidades como Educação, MAPA e outras.
- **Geração de Arquivo Excel:** Os resultados são salvos automaticamente em um arquivo Excel (.xlsx) com data e hora.
- **Logs de Execução:** Eventos como sucesso no salvamento de arquivos e erros são registrados no arquivo `calculo_repasses.log`.

## Pré-requisitos

Certifique-se de que você tem o Python 3 instalado no seu sistema. As dependências necessárias podem ser instaladas através do arquivo `requirements.txt`.

### Instalação

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

2. Geração do executável:
   Caso queira gerar um executável do script, utilize o PyInstaller:
   ```bash
   pyinstaller --onefile --icon=meu-icon1.ico calculadora-repasses.py
   ```

## Uso

1. Execute o programa:
   ```bash
   python calculadora-repasses.py
   ```

2. Escolha o tipo de cálculo que deseja realizar:
   - **1 - Prognóstico Numérico**
   - **2 - Quota Fixa**
   - **3 - Instantânea**

3. Insira o valor do GGR (Gross Gaming Revenue), e o programa calculará os valores para cada modalidade.

4. O resultado será exibido no console e salvo em um arquivo Excel.

## Manual de Usuário

As instruções para utilizar o programa estão detalhadas no manual do usuário, disponível em formato PDF neste repositório.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir uma issue ou enviar um pull request.

``` 