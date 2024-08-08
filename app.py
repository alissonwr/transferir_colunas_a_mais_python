# Importa as bibliotecas necessárias para o funcionamento do aplicativo
from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO

# Flask: framework de micro web que permite a criação de aplicações web em Python.
# request: objeto do Flask usado para lidar com dados de requisições HTTP (como formulários e uploads de arquivos).
# render_template: função que renderiza templates HTML e os envia ao navegador do usuário.
# send_file: função que permite enviar arquivos do servidor para o cliente como resposta.
# pandas (pd): biblioteca poderosa para manipulação e análise de dados, especialmente em formato de tabelas.
# BytesIO: módulo que permite manipular bytes em memória como se fossem arquivos.

# Cria uma instância do aplicativo Flask, essencial para que o aplicativo web funcione.
app = Flask(__name__)

# Define uma rota para a página inicial ('/') do aplicativo.
@app.route('/')
def index():
    # Renderiza e retorna a página HTML 'index.html' quando o usuário acessa a raiz do site.
    return render_template('index.html')

# Define a rota '/transfer' que aceita requisições POST, utilizada para processar o upload e a manipulação dos arquivos.
@app.route('/transfer', methods=['POST'])
def transfer():
    # Obtém os arquivos carregados pelo usuário no formulário da página web.
    # 'arquivo1' e 'arquivo2' são os nomes dos campos de upload no formulário.
    arquivo1 = request.files['arquivo1']
    arquivo2 = request.files['arquivo2']
    
    # Obtém os valores das colunas comuns selecionadas pelo usuário no formulário.
    # Esses valores indicam quais colunas dos arquivos Excel serão usadas para combinar os dados.
    coluna_comum_arquivo1 = request.form['coluna_comum_arquivo1']
    coluna_comum_arquivo2 = request.form['coluna_comum_arquivo2']
    
    # Lê os arquivos Excel enviados e os converte em DataFrames do pandas, uma estrutura de dados que facilita a manipulação e análise.
    # 'engine=openpyxl' é usado para garantir que a leitura do Excel seja compatível, especialmente para arquivos .xlsx.
    df1 = pd.read_excel(arquivo1, engine='openpyxl')
    df2 = pd.read_excel(arquivo2, engine='openpyxl')
    
    # Imprime as primeiras linhas de cada DataFrame no console para verificar se foram carregados corretamente.
    print("DataFrame 1 Carregado:")
    print(df1.head())  # Exibe as primeiras 5 linhas do DataFrame 1
    print("DataFrame 2 Carregado:")
    print(df2.head())  # Exibe as primeiras 5 linhas do DataFrame 2
    
    # Renomeia as colunas escolhidas pelo usuário para um nome comum ('comum') em ambos os DataFrames.
    # Isso facilita a combinação dos dados posteriormente, pois as colunas terão o mesmo nome.
    df1 = df1.rename(columns={coluna_comum_arquivo1: 'comum'})
    df2 = df2.rename(columns={coluna_comum_arquivo2: 'comum'})
    
    # Realiza uma limpeza nas colunas renomeadas:
    # - .astype(str): Converte os valores da coluna para strings.
    # - .str.strip(): Remove quaisquer espaços em branco no início ou final das strings.
    # - .str.upper(): Converte todos os caracteres para maiúsculas.
    # Isso é feito para garantir que a comparação entre as colunas seja consistente, evitando problemas de espaços ou diferenças de maiúsculas/minúsculas.
    df1['comum'] = df1['comum'].astype(str).str.strip().str.upper()
    df2['comum'] = df2['comum'].astype(str).str.strip().str.upper()
    
    # Verifica se a coluna 'comum' realmente existe em ambos os DataFrames.
    # Isso é uma verificação de segurança para garantir que a coluna necessária para a combinação está presente.
    if 'comum' not in df1.columns or 'comum' not in df2.columns:
        # Se a coluna não existir em um dos DataFrames, retorna uma mensagem de erro ao usuário.
        return "Erro: A coluna comum não existe em um dos arquivos."

    # Obtém uma lista de valores únicos (cidades) na coluna 'comum' do segundo DataFrame, excluindo valores nulos.
    # Isso servirá para filtrar os dados do primeiro DataFrame para manter apenas as cidades que estão presentes no segundo arquivo.
    cidades_comum = df2['comum'].dropna().unique()
    print("Cidades em Comum no Segundo Arquivo:")
    print(cidades_comum)  # Exibe a lista de cidades comuns no segundo DataFrame
    
    # Filtra o primeiro DataFrame para manter apenas as linhas onde a coluna 'comum' tem valores que estão na lista 'cidades_comum'.
    df1_filtrado = df1[df1['comum'].isin(cidades_comum)]
    # Filtra o segundo DataFrame para manter apenas as linhas onde a coluna 'comum' tem valores que estão na lista 'cidades_comum'.
    df2_filtrado = df2[df2['comum'].isin(cidades_comum)]
    
    # Imprime os DataFrames filtrados no console para verificação.
    print("DataFrame 1 Filtrado:")
    print(df1_filtrado)
    print("DataFrame 2 Filtrado:")
    print(df2_filtrado)

    # Verifica se o DataFrame filtrado do primeiro arquivo está vazio, o que indicaria que não há dados correspondentes.
    if df1_filtrado.empty:
        # Se não houver dados correspondentes, retorna uma mensagem de erro ao usuário.
        return "Erro: Nenhum dado correspondente encontrado no primeiro arquivo."

    # Realiza a junção (merge) dos dois DataFrames filtrados, combinando-os pela coluna 'comum'.
    # 'how="outer"' é usado para manter todas as linhas de ambos os DataFrames, mesmo se não houver correspondência perfeita.
    df_resultante = pd.merge(df1_filtrado, df2_filtrado, on='comum', how='outer')

    # Cria um buffer em memória (BytesIO) para armazenar o arquivo Excel que será gerado.
    output = BytesIO()
    
    # Cria um objeto ExcelWriter, que permitirá escrever os dados do DataFrame para um arquivo Excel.
    # O arquivo será salvo no buffer 'output' em vez de ser salvo diretamente no disco.
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Escreve o DataFrame resultante da junção em uma aba (sheet) chamada 'Dados Completos' no arquivo Excel.
        df_resultante.to_excel(writer, sheet_name='Dados Completos', index=False)
    
    # Move o cursor do buffer para o início (posição 0) para garantir que o arquivo seja lido corretamente quando for enviado.
    output.seek(0)

    # Usa a função send_file para enviar o arquivo Excel gerado para o usuário como um anexo.
    # O arquivo será baixado com o nome 'dados_completos.xlsx'.
    return send_file(output, download_name='dados_completos.xlsx', as_attachment=True)

# Condição que verifica se o script está sendo executado diretamente (e não importado como módulo).
# Se for o caso, o aplicativo Flask é iniciado em modo de depuração (debug=True).
if __name__ == '__main__':
    app.run(debug=True)
