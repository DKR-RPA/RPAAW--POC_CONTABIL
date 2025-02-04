import os
import shutil
from datetime import datetime, timedelta
import pandas as pd
import calendar


# Exemplo de uso
base_path = "./Contabilidade/"
base_dir = "./Contabilidade/"

def listar_arquivos_xlsx(diretorio):
    arquivos = []
    for root, _, files in os.walk(diretorio):  # Percorre diretórios e subdiretórios
        for f in files:
            if f.lower().strip().endswith(".xlsx"):
                arquivos.append(os.path.join(root, f))  # Salva o caminho completo
    return arquivos

def controle_pastas(base_dir):
    # Diretório base
    base_dir = os.path.expanduser(base_dir)
    poc_dir = os.path.join(base_dir, "POC")
    ano_atual = datetime.now().strftime("%Y")

    meses_portugues = {
        "January": "Janeiro", "February": "Fevereiro", "March": "Março", 
        "April": "Abril", "May": "Maio", "June": "Junho", 
        "July": "Julho", "August": "Agosto", "September": "Setembro", 
        "October": "Outubro", "November": "Novembro", "December": "Dezembro"
    }

    # Obtendo o nome do mês anterior em inglês
    mes_anterior_en = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%B")

    # Convertendo para português
    mes_anterior = meses_portugues.get(mes_anterior_en, mes_anterior_en)

    print(mes_anterior)

    # Caminhos das pastas
    ano_dir = os.path.join(poc_dir, ano_atual)
    mes_anterior_dir = os.path.join(ano_dir, mes_anterior)
    lancado_dir = os.path.join(mes_anterior_dir, "Lançado")

    # Criando as pastas, se não existirem
    for dir_path in [poc_dir, ano_dir, mes_anterior_dir, lancado_dir]:
        os.makedirs(dir_path, exist_ok=True)

    # Diretórios de saída
    diretorio_arquivos = mes_anterior_dir
    diretorio_lancado = lancado_dir

    # Verifica a existência de arquivos .xlsx na pasta POC e Mês
    arquivos_xlsx = set(listar_arquivos_xlsx(poc_dir) + listar_arquivos_xlsx(mes_anterior_dir))

    print(arquivos_xlsx)

    if not arquivos_xlsx:
        # Caso não existam, verificar na pasta do mês atual
        mes_atual = datetime.now().strftime("%B")
        mes_atual_dir = os.path.join(ano_dir, mes_atual)
        if not os.path.exists(mes_atual_dir):
            raise Exception(f"Business Exception: Nenhum arquivo .xlsx encontrado em {poc_dir} ou {mes_anterior_dir}")
    
    # Verifica se há um arquivo que contenha "POC_Composição" no nome
    arquivo_encontrado = None
    for arquivo in arquivos_xlsx:
        if "POC_Composição" in arquivo:
            arquivo_encontrado = arquivo
            print(f"O arquivo: {arquivo_encontrado}, localizado com sucesso.")
            
            # Move o arquivo encontrado para o diretório destino
            destino = os.path.join(diretorio_arquivos, os.path.basename(arquivo_encontrado))
            shutil.move(arquivo_encontrado, destino)
            break

    if not arquivo_encontrado:
        raise Exception("Business Exception: Nenhum arquivo contendo 'POC_Composição.xlsx' foi encontrado.")

    print(arquivo_encontrado)

    return True, arquivo_encontrado

def tratar_competencia(competencia):
    """
    Separa e trata o mês e o ano da variável COMPETENCIA.
    - Mês: Primeira letra maiúscula, resto minúsculo.
    - Ano: Convertido para inteiro.
    """

    # Divide a string usando '/' como separador
    partes = competencia.split('/')

    if len(partes) != 2:
        raise Exception("Business Exception: Os dados no campo Competência estão fora do formata esperado: 'Dezembro/2024'.")  # Retorna erro caso o formato esteja incorreto

    mes = partes[0].strip()  # Remove espaços extras do mês
    ano = partes[1].strip()  # Remove espaços extras do ano

    # Tratamento do mês: primeira letra maiúscula e resto minúsculo
    mes_tratado = mes.capitalize()

    # Tratamento do ano: converte para inteiro (se possível)
    try:
        ano_tratado = int(ano)
    except ValueError:
        raise Exception("Business Exception: Os dados no campo Competência fora do formata esperado: 'Dezembro/2024'.")  # Retorna erro se o ano não for um número válido

    return True, mes_tratado, ano_tratado



def processar_excel(arquivo_encontrado):
    # Lendo o arquivo Excel
    df = pd.read_excel(arquivo_encontrado, dtype=str)  # Lendo tudo como string para evitar problemas

    # Tratando os nomes das colunas (tudo maiúsculo e sem espaços)
    df.columns = [col.strip().upper().replace(" ", "_") for col in df.columns]

    # Removendo espaços extras das células e tratando INCC
    if "INCC" in df.columns:
        df["INCC"] = df["INCC"].str.strip().str.upper().map({"S": True, "N": False})

    # Tratando os demais campos removendo espaços extras nas extremidades
    for col in df.columns:
        if df[col].dtype == "object":  # Apenas para colunas de texto
            df[col] = df[col].str.strip()

    return True, df

def obter_mes_e_ultimo_dia(mes_tratado, ano_tratado):
    # Cria um dicionário de meses em português
    meses = {
    "janeiro": "01",
    "fevereiro": "02",
    "março": "03",
    "abril": "04",
    "maio": "05",
    "junho": "06",
    "julho": "07",
    "agosto": "08",
    "setembro": "09",
    "outubro": "10",
    "novembro": "11",
    "dezembro": "12"
}
    
    # Converte o mês tratado para minúsculo e encontra o número do mês
    mes_numero = meses.get(mes_tratado.lower())
    
    if mes_numero is None:
        return "Mês inválido", None, None

    # Converte mes_numero para inteiro para passar para o calendar
    mes_numero_int = int(mes_numero)
    
    # Usando calendar.monthrange para obter o último dia do mês
    _, ultimo_dia = calendar.monthrange(ano_tratado, mes_numero_int)

    # Variáveis conforme solicitado:
    # 1. inicio = "01/mes_numero/ano_tratado"
    inicio = f"01/{mes_numero}/{ano_tratado}"
    
    # 2. fim = "ultimo_dia/mes_numero/ano_tratado"
    fim = f"{ultimo_dia}/{mes_numero}/{ano_tratado}"
    
    # 3. periodo = "01-mm-aaaa"
    periodo = inicio[-7:].replace("/", "-")
    
    # 4. ultimo_periodo_fechado = data com o ultimo dia do mês
    ultimo_periodo_fechado = datetime(ano_tratado, mes_numero_int, ultimo_dia)

    ultimo_periodo_fechado = ultimo_periodo_fechado.strftime("%d-%m-%Y")
    
    return True, mes_numero, ultimo_dia, inicio, fim, periodo, ultimo_periodo_fechado



sucess, arquivo_encontrado = controle_pastas(base_dir)

sucess, df = processar_excel(arquivo_encontrado)

for idx, poc_composicao in enumerate(df.itertuples(index=True), start=1):  # O índice será acessado
    # Atribuindo valores às variáveis
    codigo = poc_composicao.CÓDIGO
    nome = poc_composicao.NOME
    diretorio = poc_composicao.DIRETÓRIO_RELATÓRIOS
    status = poc_composicao.STATUS
    competencia = poc_composicao.COMPETÊNCIA
    incc = poc_composicao.INCC
    estagio = poc_composicao.ESTÁGIO
    indice = poc_composicao.INDICE
    cadastro = poc_composicao.CADASTRO
    ajustes = poc_composicao.AJUSTES

    
    # Exibindo todos os campos junto com o número da linha
    print(f"""
    Linha: {idx}  # Linha do DataFrame
    Código: {codigo}
    Nome: {nome}
    Diretório Relatórios: {diretorio}
    Status: {status}
    Competência: {competencia}
    INCC: {incc}
    Estágio: {estagio}
    Índice: {indice}
    Cadastro: {cadastro}
    Ajustes: {ajustes}
    """)

    sucess, mes_tratado, ano_tratado = tratar_competencia(competencia)
    print(sucess)
    print(mes_tratado)
    print(ano_tratado)

    sucess, mes_numero, ultimo_dia, inicio, fim, periodo, ultimo_periodo_fechado = obter_mes_e_ultimo_dia(mes_tratado, ano_tratado)
    print(sucess)
    print(mes_numero)
    print(ultimo_dia)
    print(inicio)
    print(fim)
    print(periodo)
    print(ultimo_periodo_fechado)