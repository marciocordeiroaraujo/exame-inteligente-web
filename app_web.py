# -*- coding: utf-8 -*-
# =====================================================================
# EXAME INTELIGENTE 4.0 - VERSAO WEB (Streamlit)
# Porta do app_desktop.py para funcionar como site no navegador,
# reutilizando os mesmos arquivos JSON do desktop.
# =====================================================================
import streamlit as st
import json
import os
import random
import io
import calendar
import re
import csv
import hashlib
import unicodedata
import hmac
import shutil
from datetime import datetime, timedelta

from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.oxml import parse_xml
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import psycopg2
    _TEM_PSYCOPG2 = True
except Exception:
    psycopg2 = None
    _TEM_PSYCOPG2 = False

# =====================================================================
# 0. INICIALIZACAO: garante que rode via `streamlit run`
# (evita o modo "bare" ao executar com `python app_web.py` pelo VS Code)
# =====================================================================
import sys
if __name__ == "__main__" and not st.runtime.exists():
    try:
        subprocess_ok = True
        import subprocess
    except Exception:
        subprocess_ok = False
    print("Executando via Streamlit...")
    if subprocess_ok:
        subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
    else:
        print("Nao foi possivel iniciar o Streamlit automaticamente.")
        print('Execute manualmente no terminal: streamlit run app_web.py')
    sys.exit(0)

# =====================================================================
# 0. CONFIGURACAO INICIAL
# =====================================================================
st.set_page_config(
    page_title="Exame Inteligente 4.0 - Web",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_BANCO = "banco_questoes.json"
ARQUIVO_CONFIG = "configuracoes.json"
ARQUIVO_PLANOS = "planos_aula.json"
ARQUIVO_GRADE = "grade_horaria.json"
ARQUIVO_ANOTACOES = "anotacoes.json"
ARQUIVO_TURMAS = "turmas_alunos.json"
ARQUIVO_CONFIG_GRADE = "config_grade.json"
ARQUIVO_AVALIACOES = "avaliacoes.json"
PASTA_IMAGENS = "imagens_apoio"
ARQUIVO_USUARIOS = "usuarios.json"
PASTA_DADOS = "dados_usuarios"

if not os.path.exists(PASTA_IMAGENS):
    os.makedirs(PASTA_IMAGENS)

CONFIG_PADRAO = {
    "escola": "CEJA Profa Cecy Cialdini",
    "professor": "Marcio Cordeiro",
    "fonte": "Arial",
    "tamanho_fonte": 11,
    "margem_cm": 1.5,
    "espacamento_pt": 0,
    "usar_duas_colunas": True,
    "mostrar_aluno": True,
    "mostrar_turma": True,
    "mostrar_data": True,
    "aparencia": "System",
    "cor_tema": "blue",
    "cor_principal": "#1f538d",
    "cor_secundaria": "#14375e",
    "cor_borda_card": "",
    "cor_fundo": "#2b2b2b"
}

DESCRITORES_SAEB = [
    "D1 - Identificar a localizacao/movimentacao de objeto em mapas, croquis e outras representacoes graficas.",
    "D2 - Identificar propriedades comuns e diferencas entre figuras bidimensionais e tridimensionais, relacionando-as com as suas planificacoes.",
    "D3 - Identificar propriedades de triangulos pela comparacao de medidas de lados e angulos.",
    "D4 - Identificar relacao entre quadrilateros por meio de suas propriedades.",
    "D5 - Reconhecer a conservacao ou modificacao de medidas dos lados, do perimetro, da area em ampliacao e/ou reducao de figuras poligonais usando malhas quadriculadas.",
    "D6 - Reconhecer angulos como mudanca de direcao ou giros, identificando angulos retos e nao-retos.",
    "D7 - Reconhecer que as imagens de uma figura construida por uma transformacao homotetica sao semelhantes, identificando propriedades e/ou medidas que se modificam ou nao se alteram.",
    "D8 - Resolver problema utilizando propriedades dos poligonos (soma de seus angulos internos, numero de diagonais, calculo da medida de cada angulo interno nos poligonos regulares).",
    "D9 - Interpretar informacoes apresentadas por meio de coordenadas cartesianas.",
    "D10 - Utilizar relacoes metricas do triangulo retangulo para resolver problemas significativos.",
    "D11 - Reconhecer circulo/circunferencia, seus elementos e algumas de suas relacoes.",
    "D12 - Resolver problema envolvendo o calculo de perimetro de figuras planas.",
    "D13 - Resolver problema envolvendo o calculo de area de figuras planas.",
    "D14 - Resolver problema envolvendo nocoes de volume.",
    "D15 - Resolver problema utilizando relacoes entre diferentes unidades de medida.",
    "D16 - Identificar a localizacao de numeros inteiros na reta numerica.",
    "D17 - Identificar a localizacao de numeros racionais na reta numerica.",
    "D18 - Efetuar calculos com numeros inteiros, envolvendo as operacoes (adicao, subtracao, multiplicacao, divisao, potenciacao).",
    "D19 - Resolver problema com numeros naturais, envolvendo diferentes significados das operacoes (adicao, subtracao, multiplicacao, divisao, potenciacao).",
    "D20 - Resolver problema com numeros inteiros envolvendo as operacoes (adicao, subtracao, multiplicacao, divisao, potenciacao).",
    "D21 - Reconhecer as diferentes representacoes de um numero racional.",
    "D22 - Identificar fracao como representacao que pode estar associada a diferentes significados.",
    "D23 - Identificar fracoes equivalentes.",
    "D24 - Reconhecer as representacoes decimais dos numeros racionais como uma extensao do sistema de numeracao decimal, identificando a existencia de ordens como decimos, centesimos e milesimos.",
    "D25 - Efetuar calculos que envolvam operacoes com numeros racionais (adicao, subtracao, multiplicacao, divisao, potenciacao).",
    "D26 - Resolver problema com numeros racionais envolvendo as operacoes (adicao, subtracao, multiplicacao, divisao, potenciacao).",
    "D27 - Efetuar calculos simples com valores aproximados de radicais.",
    "D28 - Resolver problema que envolva porcentagem.",
    "D29 - Resolver problema que envolva variacao proporcional, direta ou inversa, entre grandezas.",
    "D30 - Calcular o valor numerico de uma expressao algebrica.",
    "D31 - Resolver problema que envolva equacao do 2o grau.",
    "D32 - Identificar a expressao algebrica que expressa uma regularidade observada em sequencias de numeros ou figuras (padroes).",
    "D33 - Identificar uma equacao ou inequacao do 1o grau que expressa um problema.",
    "D34 - Identificar um sistema de equacoes do 1o grau que expressa um problema.",
    "D35 - Identificar a relacao entre as representacoes algebrica e geometrica de um sistema de equacoes do 1o grau.",
    "D36 - Resolver problema envolvendo informacoes apresentadas em tabelas e/ou graficos.",
    "D37 - Associar informacoes apresentadas em listas e/ou tabelas simples aos graficos que as representam e vice-versa."
]

DESCRITORES_MAT = {
    "D1": "Localizacao/movimentacao de objeto em mapas/croquis.",
    "D2": "Propriedades e planificacoes de figuras 2D e 3D.",
    "D3": "Propriedades de triangulos (lados e angulos).",
    "D4": "Relacao entre quadrilateros.",
    "D5": "Conservacao/modificacao em figuras poligonais (malhas).",
    "D6": "Angulos: mudanca de direcao, giros, retos e nao retos.",
    "D7": "Transformacao homotetica e semelhanca.",
    "D8": "Propriedades dos poligonos (soma de angulos, diagonais).",
    "D9": "Informacoes apresentadas em coordenadas cartesianas.",
    "D10": "Relacoes metricas do triangulo retangulo.",
    "D11": "Circulo/circunferencia e suas relacoes.",
    "D12": "Calculo de perimetro de figuras planas.",
    "D13": "Calculo de area de figuras planas.",
    "D14": "Nocoes de volume.",
    "D15": "Relacoes entre diferentes unidades de medida.",
    "D16": "Numeros inteiros na reta numerica.",
    "D17": "Numeros racionais na reta numerica.",
    "D18": "Calculos com numeros inteiros (4 operacoes).",
    "D19": "Problemas com numeros naturais (4 operacoes e potencias).",
    "D20": "Problemas com numeros inteiros (4 operacoes e potencias).",
    "D21": "Diferentes representacoes de um numero racional.",
    "D22": "Fracao com diferentes significados.",
    "D23": "Fracoes equivalentes.",
    "D24": "Representacoes decimais (decimos, centesimos, milesimos).",
    "D25": "Calculos com numeros racionais.",
    "D26": "Problemas com numeros racionais.",
    "D27": "Calculos simples com valores aproximados de radicais.",
    "D28": "Problemas com porcentagem.",
    "D29": "Variacao proporcional, direta ou inversa.",
    "D30": "Valor numerico de uma expressao algebrica.",
    "D31": "Equacao do 2o grau em problemas.",
    "D32": "Expressao algebrica em sequencias numericas (padroes).",
    "D33": "Equacao ou inequacao do 1o grau num problema.",
    "D34": "Sistema de equacoes do 1o grau num problema.",
    "D35": "Representacoes algebrica e geometrica de sistema do 1o grau.",
    "D36": "Informacoes apresentadas em tabelas e/ou graficos.",
    "D37": "Associacao entre listas/tabelas e graficos."
}

def esc(s):
    import html
    return html.escape(str(s), quote=True)

# =====================================================================
# 0A. BANCO EXTERNO (Supabase/Postgres) - persistencia na nuvem
#     Quando BANCO_URL estiver configurado (.streamlit/secrets.toml ou
#     variavel de ambiente), todos os dados passam a viver no Postgres.
#     As funcoes carregar_* / salvar_* continuam iguais para o resto do app.
# =====================================================================
def _obter_banco_url():
    env = os.environ.get("BANCO_URL", "").strip()
    if env:
        if env.lower() in ("0", "off", "false", "nenhum"):
            return ""
        return env
    try:
        if "BANCO_URL" in st.secrets:
            return str(st.secrets["BANCO_URL"]).strip()
    except Exception:
        pass
    return ""

BANCO_ATIVO = bool(_obter_banco_url()) and _TEM_PSYCOPG2

def _conectar_banco():
    url = _obter_banco_url()
    if "?" in url:
        url = url.split("?", 1)[0]
    conn = psycopg2.connect(url, sslmode="require")
    conn.autocommit = True
    return conn

def _com_conexao(fn):
    conn = None
    try:
        conn = _conectar_banco()
        with conn.cursor() as cur:
            return fn(conn, cur)
    except Exception:
        raise
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass

_TABELAS_OK = False

def _garantir_tabelas():
    global _TABELAS_OK
    if _TABELAS_OK:
        return
    def _criar(conn, cur):
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_dados (
                chave text PRIMARY KEY,
                valor jsonb NOT NULL,
                atualizado_em timestamptz NOT NULL DEFAULT now()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_imagens (
                chave text PRIMARY KEY,
                dados bytea NOT NULL,
                atualizado_em timestamptz NOT NULL DEFAULT now()
            )
        """)
    _com_conexao(_criar)
    _TABELAS_OK = True

def db_get_json(chave):
    try:
        _garantir_tabelas()
        def _get(conn, cur):
            cur.execute("SELECT valor FROM app_dados WHERE chave=%s", (chave,))
            linha = cur.fetchone()
            return linha[0] if linha else None
        return _com_conexao(_get)
    except Exception:
        return None

def db_set_json(chave, dados):
    try:
        _garantir_tabelas()
        texto = json.dumps(dados, ensure_ascii=False)
        def _set(conn, cur):
            cur.execute(
                "INSERT INTO app_dados (chave, valor, atualizado_em) "
                "VALUES (%s, %s::jsonb, now()) "
                "ON CONFLICT (chave) DO UPDATE "
                "SET valor=EXCLUDED.valor, atualizado_em=now()",
                (chave, texto))
        _com_conexao(_set)
    except Exception as e:
        print("[BANCO] Erro ao salvar", chave, ":", e)

def db_get_bytes(chave):
    try:
        _garantir_tabelas()
        def _get(conn, cur):
            cur.execute("SELECT dados FROM app_imagens WHERE chave=%s", (chave,))
            linha = cur.fetchone()
            return bytes(linha[0]) if linha else None
        return _com_conexao(_get)
    except Exception:
        return None

def db_set_bytes(chave, dados):
    try:
        _garantir_tabelas()
        def _set(conn, cur):
            cur.execute(
                "INSERT INTO app_imagens (chave, dados, atualizado_em) "
                "VALUES (%s, %s, now()) "
                "ON CONFLICT (chave) DO UPDATE "
                "SET dados=EXCLUDED.dados, atualizado_em=now()",
                (chave, psycopg2.Binary(bytes(dados))))
        _com_conexao(_set)
    except Exception as e:
        print("[BANCO] Erro ao salvar imagem", chave, ":", e)
    except Exception as e:
        print("[BANCO] Erro ao salvar imagem", chave, ":", e)

def _chave_de_caminho(caminho):
    if not caminho:
        return None
    norm = str(caminho).replace("\\", "/")
    partes = [p for p in norm.split("/") if p]
    if partes == [ARQUIVO_USUARIOS]:
        return "g::" + ARQUIVO_USUARIOS
    if len(partes) >= 3 and partes[0] == PASTA_DADOS:
        return "u::" + partes[1] + "::" + "/".join(partes[2:])
    return None

def imagem_existe(caminho):
    chave = _chave_de_caminho(caminho)
    if BANCO_ATIVO and chave:
        return db_get_bytes(chave) is not None
    return bool(caminho) and os.path.exists(caminho)

def imagem_bytes(caminho):
    chave = _chave_de_caminho(caminho)
    if BANCO_ATIVO and chave:
        dados = db_get_bytes(chave)
        if dados is not None:
            return dados
    if caminho and os.path.exists(caminho):
        with open(caminho, "rb") as f:
            return f.read()
    return b""

def salvar_imagem_usuario(nome_arquivo, dados):
    caminho = os.path.join(pasta_imagens(), nome_arquivo)
    chave = _chave_de_caminho(caminho)
    if BANCO_ATIVO and chave:
        db_set_bytes(chave, dados)
    else:
        with open(caminho, "wb") as f:
            f.write(bytes(dados))
    return caminho

def _testar_banco():
    try:
        _garantir_tabelas()
        return True
    except Exception as e:
        print("[BANCO] Nao foi possivel conectar:", e)
        return False

if BANCO_ATIVO and not _testar_banco():
    print("[BANCO] Desativando banco externo - usando arquivos locais.")
    BANCO_ATIVO = False

# =====================================================================
# 1. UTILITARIOS DE DADOS (JSON) - mesmos arquivos do desktop
#    Desde o port web, cada conta tem a propria pasta em dados_usuarios/
# =====================================================================
def usuario_atual():
    return st.session_state.get("usuario", None)

def pasta_usuario(usuario=None):
    usuario = usuario or usuario_atual()
    if not usuario:
        return None
    pasta = os.path.join(PASTA_DADOS, str(usuario))
    os.makedirs(pasta, exist_ok=True)
    return pasta

def caminho_usuario(nome_arquivo):
    pasta = pasta_usuario()
    if not pasta:
        return None
    return os.path.join(pasta, nome_arquivo)

def pasta_imagens():
    pasta = pasta_usuario()
    if not pasta:
        return PASTA_IMAGENS
    destino = os.path.join(pasta, PASTA_IMAGENS)
    os.makedirs(destino, exist_ok=True)
    return destino

def carregar_json(caminho, padrao):
    if BANCO_ATIVO:
        chave = _chave_de_caminho(caminho)
        if chave:
            dados = db_get_json(chave)
            if dados is not None:
                return dados
    if caminho and os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as arquivo:
                return json.load(arquivo)
        except Exception:
            return padrao
    return padrao

def salvar_json(caminho, dados):
    if not caminho:
        return
    if BANCO_ATIVO:
        chave = _chave_de_caminho(caminho)
        if chave:
            db_set_json(chave, dados)
            return
    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=4)

# ---- Contas de usuario (login local em JSON) ----
def carregar_usuarios():
    dados = carregar_json(ARQUIVO_USUARIOS, [])
    return dados if isinstance(dados, list) else []

def salvar_usuarios(usuarios):
    salvar_json(ARQUIVO_USUARIOS, usuarios)

def hash_senha(senha, salt_hex=None):
    salt = bytes.fromhex(salt_hex) if salt_hex else os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt, 200000)
    return salt.hex(), digest.hex()

def usuario_existe(usuario):
    alvo = (usuario or "").lower()
    return any(u.get("usuario", "").lower() == alvo for u in carregar_usuarios())

def autenticar_usuario(usuario, senha):
    alvo = (usuario or "").strip().lower()
    for u in carregar_usuarios():
        if u.get("usuario", "").lower() == alvo:
            salt, digest = u.get("senha_hash", ["", ""])
            if not salt:
                return None
            _, novo_digest = hash_senha(senha, salt)
            if hmac.compare_digest(novo_digest, digest):
                return u
            return None
    return None

def conta_atual(usuario=None):
    usuario = usuario or usuario_atual()
    if not usuario:
        return None
    for u in carregar_usuarios():
        if u.get("usuario", "").lower() == usuario.lower():
            return u
    return None

def migrar_dados_primeira_conta(usuario):
    pasta = pasta_usuario(usuario)
    if not pasta:
        return
    for nome_arquivo in [ARQUIVO_BANCO, ARQUIVO_CONFIG, ARQUIVO_PLANOS,
                         ARQUIVO_GRADE, ARQUIVO_ANOTACOES, ARQUIVO_TURMAS,
                         ARQUIVO_CONFIG_GRADE, ARQUIVO_AVALIACOES]:
        origem = nome_arquivo
        destino = os.path.join(pasta, nome_arquivo)
        if os.path.exists(origem) and not os.path.exists(destino):
            try:
                shutil.copy2(origem, destino)
            except Exception:
                pass
    origem_imgs = PASTA_IMAGENS
    destino_imgs = os.path.join(pasta, PASTA_IMAGENS)
    if os.path.isdir(origem_imgs) and not os.path.isdir(destino_imgs):
        try:
            shutil.copytree(origem_imgs, destino_imgs)
        except Exception:
            pass
    _ajustar_imagens_banco(pasta)

def _ajustar_imagens_banco(pasta):
    caminho_banco = os.path.join(pasta, ARQUIVO_BANCO)
    if not os.path.exists(caminho_banco):
        return
    try:
        with open(caminho_banco, "r", encoding="utf-8") as f:
            banco = json.load(f)
        prefixo = os.path.join(PASTA_DADOS, os.path.basename(pasta), PASTA_IMAGENS)
        mudou = False
        for q in banco:
            img = q.get("imagem", "")
            if img and img.startswith(PASTA_IMAGENS + os.sep):
                q["imagem"] = os.path.join(prefixo, os.path.basename(img))
                mudou = True
        if mudou:
            salvar_json(caminho_banco, banco)
    except Exception:
        pass

def criar_conta(nome, usuario, senha):
    usuarios = carregar_usuarios()
    primeira = len(usuarios) == 0
    salt, digest = hash_senha(senha)
    usuarios.append({
        "nome": nome.strip(),
        "usuario": usuario.strip(),
        "senha_hash": [salt, digest],
        "cookie_token": os.urandom(16).hex(),
        "onboarding_pendente": bool(not nome.strip()),
        "criado_em": datetime.now().strftime("%d/%m/%Y %H:%M")
    })
    salvar_usuarios(usuarios)
    if primeira:
        migrar_dados_primeira_conta(usuario.strip())
    return usuario.strip()

def conta_precisa_onboarding():
    conta = conta_atual()
    return bool(conta and conta.get("onboarding_pendente"))

def concluir_onboarding(nome, escola):
    conta = conta_atual()
    if conta:
        usuarios = carregar_usuarios()
        alvo = str(conta.get("usuario", "")).lower()
        for u in usuarios:
            if str(u.get("usuario", "")).lower() == alvo:
                u["nome"] = (nome or "").strip()
                u["onboarding_pendente"] = False
                break
        salvar_usuarios(usuarios)
    config = carregar_config()
    config["professor"] = (nome or "").strip()
    config["escola"] = (escola or "").strip()
    salvar_config(config)

# ---- Lembrar da conta (cookie) ----
EI_COOKIE = "ei_usuario"

def garantir_cookie_token():
    conta = conta_atual()
    if not conta:
        return None, None
    if not conta.get("cookie_token"):
        usuarios = carregar_usuarios()
        alvo = str(conta.get("usuario", "")).lower()
        for u in usuarios:
            if str(u.get("usuario", "")).lower() == alvo:
                u["cookie_token"] = os.urandom(16).hex()
                break
        salvar_usuarios(usuarios)
        conta = conta_atual()
    return conta.get("usuario", ""), conta.get("cookie_token", "")

def _autologin_por_cookie():
    if st.session_state.get("deslogado_manual"):
        return False
    try:
        valor = st.context.cookies.get(EI_COOKIE, "")
    except Exception:
        return False
    if not isinstance(valor, str) or not valor or "|" not in valor:
        return False
    usuario, token = valor.split("|", 1)
    conta = conta_atual(usuario)
    if conta and conta.get("cookie_token") and hmac.compare_digest(str(conta.get("cookie_token")), token):
        st.session_state["usuario"] = conta["usuario"]
        return True
    return False

def garantir_usuario_teste():
    if not usuario_existe("_teste_"):
        criar_conta("Usuario de Teste", "_teste_", "teste123")
    return "_teste_"

# ---- Acesso por entidade (sempre na pasta da conta logada) ----
def carregar_banco():
    return carregar_json(caminho_usuario(ARQUIVO_BANCO), [])

def carregar_config():
    padrao = dict(CONFIG_PADRAO)
    conta = conta_atual()
    if conta and conta.get("nome"):
        padrao["professor"] = conta["nome"]
    else:
        padrao["professor"] = "Professor"
    padrao["escola"] = ""
    return carregar_json(caminho_usuario(ARQUIVO_CONFIG), padrao)

def carregar_planos():
    return carregar_json(caminho_usuario(ARQUIVO_PLANOS), {})

def _norm_dia(dia):
    return unicodedata.normalize("NFD", str(dia)).encode("ascii", "ignore").decode("ascii")

def carregar_grade():
    grade = carregar_json(caminho_usuario(ARQUIVO_GRADE), [])
    for item in grade:
        item["dia"] = _norm_dia(item.get("dia", ""))
    return grade

def carregar_anotacoes():
    return carregar_json(caminho_usuario(ARQUIVO_ANOTACOES), [])

def carregar_turmas():
    return carregar_json(caminho_usuario(ARQUIVO_TURMAS), {})

def carregar_config_grade():
    try:
        return int(carregar_json(caminho_usuario(ARQUIVO_CONFIG_GRADE), {}).get("max_aulas", 6))
    except Exception:
        return 6

def carregar_avaliacoes():
    return carregar_json(caminho_usuario(ARQUIVO_AVALIACOES), [])

def salvar_banco(dados):
    salvar_json(caminho_usuario(ARQUIVO_BANCO), dados)

def salvar_config(dados):
    salvar_json(caminho_usuario(ARQUIVO_CONFIG), dados)

def salvar_planos(dados):
    salvar_json(caminho_usuario(ARQUIVO_PLANOS), dados)

def salvar_grade(dados):
    for item in dados:
        item["dia"] = _norm_dia(item.get("dia", ""))
    salvar_json(caminho_usuario(ARQUIVO_GRADE), dados)

def salvar_anotacoes(dados):
    salvar_json(caminho_usuario(ARQUIVO_ANOTACOES), dados)

def salvar_turmas(dados):
    salvar_json(caminho_usuario(ARQUIVO_TURMAS), dados)

def salvar_config_grade(max_aulas):
    salvar_json(caminho_usuario(ARQUIVO_CONFIG_GRADE), {"max_aulas": max_aulas})

def salvar_avaliacoes(dados):
    salvar_json(caminho_usuario(ARQUIVO_AVALIACOES), dados)

def primeiro_nome_professor(config):
    nome = (config.get("professor") or "").strip() or "Professor"
    return nome.split()[0]

MESES_PT = ["", "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
DIAS_COMPLETOS = ["Segunda-feira", "Terca-feira", "Quarta-feira",
                  "Quinta-feira", "Sexta-feira", "Sabado", "Domingo"]
DIAS_UTEIS = ["Segunda-feira", "Terca-feira", "Quarta-feira",
              "Quinta-feira", "Sexta-feira"]
DIAS_CURTO = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
# -*- coding: utf-8 -*-
LOGO_ABA_B64 = "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"
LOGO_INICIO_B64 = "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"

# =====================================================================
# 2. TEMA / CSS (replica a aparencia do CustomTkinter)
# =====================================================================
def cor_texto_legivel(hex_cor):
    try:
        h = hex_cor.lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        luminancia = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return "#0d0d0d" if luminancia > 0.62 else "#ffffff"
    except Exception:
        return "#ffffff"

def ajustar_cor(hex_cor, fator):
    try:
        h = hex_cor.lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        r = max(0, min(255, int(r * fator)))
        g = max(0, min(255, int(g * fator)))
        b = max(0, min(255, int(b * fator)))
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return hex_cor

def cor_rgba(hex_cor, alpha):
    try:
        h = hex_cor.lstrip("#")
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        return f"rgba({r},{g},{b},{alpha})"
    except Exception:
        return hex_cor

CSS_TEMPLATE = """
<style>
:root {
    color-scheme: @@SCHEME@@;
    --cor-p: @@CORP@@;
    --cor-s: @@CORS@@;
    --cor-ph: @@CORP_HOVER@@;
    --cor-pd: @@CORP_DEEP@@;
    --cor-texto: @@TEXTO@@;
    --cor-cinza: @@CINZA@@;
    --card-bg: @@CARDBG@@;
    --borda: @@BORDA@@;
    --fundo: @@FUNDO@@;
    --btn-fg: @@BTN@@;
}
html, body, .stApp {
    background: var(--fundo) !important;
    color: var(--cor-texto) !important;
    font-family: 'Segoe UI', system-ui, -apple-system, 'Roboto', Arial, sans-serif;
}
#MainMenu, footer {visibility: hidden;}
header[data-testid="stHeader"] {
    background: transparent !important;
    border-bottom: 1px solid transparent;
}
[data-testid="stDecoration"] {display: none !important;}
.block-container {padding-top: 0.8rem; padding-bottom: 2.5rem; max-width: 1440px;}
h1, h2, h3, h4, h5 {color: var(--cor-texto) !important; letter-spacing: -0.01em;}
p, label, .stMarkdown, .stTextInput input, .stTextArea textarea,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div,
.stNumberInput input, .stDateInput input, .stTimeInput input {color: var(--cor-texto) !important;}
.stApp a {color: #17a2b8; text-decoration: none;}
.stApp a:hover {text-decoration: underline;}

/* ---------------- Scrollbar ---------------- */
::-webkit-scrollbar {width: 10px; height: 10px;}
::-webkit-scrollbar-thumb {background: @@SCROLL@@; border-radius: 8px;}
::-webkit-scrollbar-track {background: transparent;}

/* ---------------- Sidebar ---------------- */
section[data-testid="stSidebar"] {
    background: linear-gradient(185deg, var(--cor-s) 0%, @@CORS_GRAD@@ 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.10);
}
@media (min-width: 769px) {
    section[data-testid="stSidebar"] {width: 16.5rem !important; min-width: 16.5rem !important;}
}
section[data-testid="stSidebar"] {color: #ececec !important;}
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    padding-top: 2px !important;
    padding-bottom: 6px !important;
}
@media (min-width: 769px) {
    section[data-testid="stSidebar"] [data-testid="stSidebarHeader"] {
        display: none !important;
    }
}
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
section[data-testid="stSidebar"] .stCaption {
    color: #ececec !important;
}
.logo-web {
    font-size: 1.85rem; font-weight: 800; line-height: 1.1; text-align: center;
    padding: 2px 8px 4px 8px; letter-spacing: -0.02em;
    background: linear-gradient(90deg, #ffffff, #e6c8ff);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
}
.logo-web small {font-size: .5em; letter-spacing: 4px; font-weight: 700;
    -webkit-text-fill-color: rgba(255,255,255,.70); color: rgba(255,255,255,.70);}
.logo-aba {
    text-align: center; padding: 10px 0 4px 0; line-height: 0;
}
.logo-aba img {
    width: 152px; height: auto; display: inline-block;
    filter: drop-shadow(0 6px 16px rgba(0,0,0,.28));
}
.nav-secao {
    font-size: .66rem; font-weight: 800; letter-spacing: 2.5px; color: rgba(255,255,255,.55) !important;
    padding: 10px 12px 4px 12px; text-transform: uppercase;
    border-top: 1px solid rgba(255,255,255,.08); margin-top: 2px;
}
section[data-testid="stSidebar"] .stButton > button {
    width: 100%; text-align: left; justify-content: flex-start; gap: 10px;
    background: transparent; border: 1px solid transparent; color: rgba(255,255,255,.85) !important;
    border-radius: 9px; padding: .40rem .68rem; font-size: .87rem; font-weight: 600;
    margin-bottom: 2px; transition: all .18s ease; box-shadow: none;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12) !important; color: #fff !important;
    transform: translateX(3px); box-shadow: none;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: var(--cor-p) !important; color: var(--btn-fg) !important; font-weight: 700;
    border: 1px solid rgba(255,255,255,.28);
    box-shadow: 0 3px 12px rgba(0,0,0,.28);
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] * {
    color: var(--btn-fg) !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
    background: var(--cor-ph) !important; transform: none;
}
section[data-testid="stSidebar"] .stButton > button[kind="secondary"] * {
    color: rgba(255,255,255,.85) !important;
}

/* ---------------- Botoes ---------------- */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
    border-radius: 10px; font-weight: 600; letter-spacing: .01em;
    transition: all .18s ease; box-shadow: 0 1px 3px rgba(0,0,0,.08);
}
.stButton > button:hover, .stDownloadButton > button:hover {
    box-shadow: 0 4px 14px rgba(0,0,0,.16); transform: translateY(-1px);
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"],
.stFormSubmitButton > button[kind="primary"] {
    background: linear-gradient(180deg, var(--cor-p) 0%, var(--cor-pd) 100%);
    color: var(--btn-fg); border: 1px solid @@CORP_EDGE@@;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover,
.stFormSubmitButton > button[kind="primary"]:hover {
    background: linear-gradient(180deg, var(--cor-ph) 0%, var(--cor-pd) 100%);
    box-shadow: 0 6px 18px @@CORP_SHADOW@@;
}
.stButton > button[kind="secondary"] {
    background: var(--card-bg); color: var(--cor-texto); border: 1px solid var(--borda);
}
.stButton > button[kind="secondary"]:hover {background: @@CORS_SOFT@@;}
.stButton > button:focus-visible, .stDownloadButton > button:focus-visible,
.stFormSubmitButton > button:focus-visible {outline: 3px solid @@CORP_SOFT@@; outline-offset: 1px;}

/* ---------------- Inputs ---------------- */
.stTextInput input, .stTextArea textarea, .stNumberInput input,
.stDateInput input, .stTimeInput input,
div[data-baseweb="input"], div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
div[data-baseweb="select"] input,
.stMultiSelect div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] input,
[data-testid="stColorPicker"] input {
    background: var(--card-bg) !important;
    color: var(--cor-texto) !important;
    border-radius: 10px !important; border-color: var(--borda) !important;
    transition: border-color .15s ease, box-shadow .15s ease;
}
.stTextInput input::placeholder, .stTextArea textarea::placeholder {color: var(--cor-cinza) !important; opacity: 1;}
.stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus,
div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within {
    border-color: @@CORS@@ !important; box-shadow: 0 0 0 3px @@CORP_SOFT@@ !important;
}
/* dropdown / menu */
[data-baseweb="popover"] [role="listbox"], [data-baseweb="menu"], [data-baseweb="popover"] ul {
    background: var(--card-bg) !important; color: var(--cor-texto) !important;
}
[data-baseweb="popover"] li {color: var(--cor-texto) !important;}
[data-baseweb="popover"] li:hover, [data-baseweb="popover"] li[aria-selected="true"] {
    background: @@CORS_SOFT@@ !important; color: var(--cor-texto) !important;
}
[data-baseweb="select"] span, [data-baseweb="select"] div {color: var(--cor-texto) !important;}
/* labels */
[data-testid="stCheckbox"] label p, [data-testid="stRadio"] label p,
[data-testid="stSelectbox"] label, [data-testid="stMultiSelect"] label,
[data-testid="stNumberInput"] label, [data-testid="stTextInput"] label,
[data-testid="stTextArea"] label, [data-testid="stDateInput"] label,
[data-testid="stTimeInput"] label, [data-testid="stSlider"] label,
[data-testid="stFileUploader"] label {
    color: var(--cor-texto) !important;
}
/* multiselect tags / slider / file uploader / expander / tabela */
[data-testid="stMultiSelect"] span[data-baseweb="tag"] {
    background: @@CORS_SOFT@@ !important; color: var(--cor-texto) !important;
}
[data-baseweb="slider"] div[role="slider"] {background: var(--cor-p) !important; border-color: var(--cor-p) !important;}
[data-testid="stFileUploaderDropzone"] {
    background: var(--card-bg) !important; border: 1px dashed var(--borda) !important;
    border-radius: 10px; color: var(--cor-texto) !important;
}
[data-testid="stFileUploaderDropzone"] span, [data-testid="stFileUploaderDropzone"] small {
    color: var(--cor-texto) !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background: var(--cor-p) !important;
    color: var(--btn-fg) !important;
    border: 1px solid @@CORP_EDGE@@ !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    box-shadow: 0 1px 3px rgba(0,0,0,.08);
}
[data-testid="stFileUploaderDropzone"] button:hover {
    background: var(--cor-ph) !important;
}
[data-testid="stExpander"] {border: 1px solid var(--borda); border-radius: 12px;}
[data-testid="stExpander"] summary {color: var(--cor-texto) !important;}
[data-testid="stTable"] {background: var(--card-bg); border-radius: 10px; overflow: hidden;}

/* ---------------- Cards ---------------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--card-bg); border: 1px solid var(--borda) !important; border-radius: 14px;
    padding: 4px 12px; box-shadow: 0 1px 3px rgba(0,0,0,.05);
    transition: box-shadow .2s ease, border-color .2s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 6px 20px rgba(0,0,0,.10); border-color: @@CORP_SOFT@@ !important;
}
.card {
    background: var(--card-bg); border: 1px solid var(--borda); border-radius: 14px;
    padding: 14px 16px; margin-bottom: 12px;
    border-top: 3px solid @@COR_BORDA@@;
    box-shadow: 0 1px 3px rgba(0,0,0,.05); transition: box-shadow .2s ease, transform .2s ease;
}
.card:hover {box-shadow: 0 6px 20px rgba(0,0,0,.10); transform: translateY(-1px);}
/* Cards via st.container(key="card_...") -> classe st-key-card_* */
div[class*="st-key-card_"] {
    background: var(--card-bg); border: 1px solid var(--borda); border-radius: 14px;
    padding: 14px 16px; margin-bottom: 12px;
    border-top: 3px solid @@COR_BORDA@@;
    box-shadow: 0 1px 3px rgba(0,0,0,.05); transition: box-shadow .2s ease;
}
div[class*="st-key-card_"]:hover {box-shadow: 0 6px 20px rgba(0,0,0,.10);}

/* Dashboard (somente telas >= 769px): grade que preenche o espaco visivel,
   com borda superior colorida em cada campo e fundo branco puro. */
@media (min-width: 769px) {
    div[class*="st-key-dash_wrap"] {
        display: grid !important;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1.6fr);
        grid-template-rows: 1fr 1fr auto;
        grid-template-areas:
            "mural cal"
            "grade cal"
            "agenda agenda";
        gap: 12px;
        align-items: stretch;
        min-height: calc(100vh - 235px);
    }
    div[class*="st-key-dash_wrap"] > div { min-width: 0; }
    div[class*="st-key-dash_wrap"] > div:has(> div.stVerticalBlock.st-key-card_dash_mural)  { grid-area: mural; }
    div[class*="st-key-dash_wrap"] > div:has(> div.stVerticalBlock.st-key-card_dash_grade)  { grid-area: grade; }
    div[class*="st-key-dash_wrap"] > div:has(> div.stVerticalBlock.st-key-card_dash_cal)    { grid-area: cal; }
    div[class*="st-key-dash_wrap"] > div:has(> div.stVerticalBlock.st-key-card_dash_agenda) { grid-area: agenda; }
    div[class*="st-key-dash_wrap"] > div > div.stVerticalBlock { height: 100%; }
    div[class*="st-key-dash_wrap"] div[class*="st-key-card_"] { margin-bottom: 0 !important; }
    div[class*="st-key-card_dash_agenda"] { min-height: 150px; }
}
.card-titulo {font-weight: 800; font-size: .95rem; margin-bottom: 4px; color: var(--cor-texto);}
.card-sub {font-size: .82rem; color: var(--cor-cinza);}
.tag {
    display: inline-block; background: @@CORS_SOFT@@; color: @@TAGS@@;
    border-radius: 20px; padding: 3px 12px; font-size: .72rem; font-weight: 800; letter-spacing: .02em;
}
.postit {
    border-radius: 10px; padding: 12px 14px; margin-bottom: 10px;
    border: 1px solid rgba(0,0,0,0.12); box-shadow: 0 2px 6px rgba(0,0,0,0.10);
    transition: transform .2s ease, box-shadow .2s ease;
}
.postit:hover {transform: translateY(-2px) rotate(-.4deg); box-shadow: 0 8px 20px rgba(0,0,0,.18);}
.postit .pt-titulo {font-weight: 800; font-size: .95rem; color: inherit; margin-bottom: 4px;}
.postit .pt-tag {font-size: .72rem; color: inherit; opacity: .78; margin-bottom: 4px;}
.postit .pt-conteudo {font-size: .86rem; color: inherit; white-space: pre-wrap;}

/* ---------------- Calendario ---------------- */
.cal-table {border-collapse: separate; border-spacing: 4px; width: 100%;}
.cal-table th {font-size: .7rem; text-transform: uppercase; letter-spacing: .08em; color: var(--cor-cinza); padding: 4px; text-align: center;}
.cal-cell {border: 1px solid var(--borda); text-align: center; font-size: .8rem; font-weight: 700;
    border-radius: 10px; width: 13%; padding: 8px 0; color: var(--cor-texto); background: var(--card-bg);
    transition: transform .15s ease, box-shadow .15s ease;}
.cal-link {text-decoration: none; display: block; padding: 8px 0; color: inherit !important; border-radius: 10px;}
.cal-hoje {background: @@CORP@@; color: @@BTN@@ !important; box-shadow: 0 3px 10px @@CORP_SHADOW@@;}
.cal-aula {background: #28a745; color: #fff !important;}
.cal-link:hover {background: #17a2b8; color: #fff !important; transform: scale(1.05);
    box-shadow: 0 4px 12px rgba(0,0,0,.20);}
.cal-scroll {overflow-x: auto; -webkit-overflow-scrolling: touch;}

/* Calendario com widgets (Dashboard): escopo nos botoes de dia.
   Nunca quebrar um dia de dois digitos em duas linhas: bloqueia quebra
   de palavra (overflow-wrap/word-break) e, no mobile, reduz a fonte. */
div[class*="st-key-cal_d_"] button {
    padding: 0.2rem 0.1rem !important;
    min-height: 2.05rem;
    font-weight: 700;
    border-radius: 10px;
    white-space: nowrap !important;
    overflow-wrap: normal !important;
    word-break: normal !important;
    overflow: hidden;
}
div[class*="st-key-cal_d_"] button [data-testid="stMarkdownContainer"],
div[class*="st-key-cal_d_"] button [data-testid="stMarkdownContainer"] p {
    white-space: nowrap !important;
    overflow-wrap: normal !important;
    word-break: normal !important;
    margin: 0 !important;
}

/* ---------------- Mini grade (Grade Semanal no Dashboard) ---------------- */
.gmini {display: grid; grid-template-columns: repeat(auto-fill, minmax(108px, 1fr)); gap: 6px;}
.gmini-dia {background: @@MINI_DIA@@; border: 1px solid var(--borda); border-radius: 10px; padding: 6px; min-width: 0;}
.gmini-nome {font-size: .72rem; font-weight: 700; color: @@CORS@@; margin-bottom: 4px; text-transform: uppercase; letter-spacing: .05em;}
.gmini-aula {background: @@MINI_AULA@@; border-radius: 6px; padding: 4px 8px; margin-bottom: 4px;
    font-size: .72rem; color: @@MINI_TXT@@; border: 1px solid transparent;}
.gmini-aula:hover {border-color: @@CORS@@;}

/* ---------------- Grade Semanal / Agenda ---------------- */
.gcell, .dash-item {
    background: var(--card-bg); border: 1px solid var(--borda); border-radius: 10px;
    padding: 8px 10px; margin-bottom: 6px;
    transition: border-color .15s ease, box-shadow .15s ease;
}
.gcell {position: relative;}
.gcell:hover, .dash-item:hover {border-color: @@CORS@@; box-shadow: 0 3px 10px rgba(0,0,0,.08);}
.gcell-aula {font-size: .72rem; font-weight: 700; color: @@CORS@@;}
.gcell-turma {font-weight: 700; font-size: .9rem; color: var(--cor-texto);}
.gcell-disc {font-size: .75rem; color: var(--cor-cinza);}
.gcell-del {
    position: absolute; top: 3px; left: 5px; z-index: 5;
    display: none; color: #e74c3c !important; font-weight: 800; font-size: .85rem;
    line-height: 1; text-decoration: none !important; padding: 2px 5px; border-radius: 6px;
    background: rgba(255,255,255,0.85);
}
.gcell:hover .gcell-del {display: block;}
.gcell-del:hover {color: #fff !important; background: #e74c3c !important;}
.dash-item-titulo {font-weight: 700; font-size: .85rem; color: var(--cor-texto);}
.dash-item-sub {font-size: .8rem; color: var(--cor-cinza);}
.pend-item {
    background: rgba(220,53,69,0.10); border: 1px solid rgba(220,53,69,0.45);
    border-radius: 10px; padding: 8px 12px; margin-bottom: 6px;
    color: var(--cor-texto); font-size: .85rem;
}

/* ---------------- Metricas ---------------- */
div[data-testid="stMetric"] {
    background: var(--card-bg); border: 1px solid var(--borda); border-radius: 14px; padding: 12px 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,.05);
}
div[data-testid="stMetricLabel"] {font-weight: 700 !important; color: var(--cor-cinza) !important;}
div[data-testid="stMetricValue"] {color: @@CORS@@ !important; font-size: 1.7rem !important; font-weight: 800 !important;}

/* ---------------- Abas ---------------- */
.stTabs [data-baseweb="tab-list"] {gap: 4px; border-bottom: 1px solid var(--borda);}
.stTabs [data-baseweb="tab"] {border-radius: 10px 10px 0 0; padding: 9px 18px; font-weight: 600;
    color: var(--cor-cinza) !important; transition: all .15s ease;}
.stTabs [data-baseweb="tab"]:hover {color: var(--cor-texto) !important; background: @@CORS_SOFT@@;}
.stTabs [data-baseweb="tab"][aria-selected="true"] {color: @@CORS@@ !important; font-weight: 800;
    border-bottom: 3px solid @@CORS@@;}

/* ---------------- Dialogs / alerts / dataframe ---------------- */
[data-testid="stDialog"] {
    border-radius: 16px; box-shadow: 0 24px 60px rgba(0,0,0,.35);
    background: var(--card-bg) !important; color: var(--cor-texto) !important;
}
[data-testid="stDialog"] p, [data-testid="stDialog"] label,
[data-testid="stDialog"] .stMarkdown, [data-testid="stDialog"] h1,
[data-testid="stDialog"] h2, [data-testid="stDialog"] h3,
[data-testid="stDialog"] h4, [data-testid="stDialog"] h5 {
    color: var(--cor-texto) !important;
}
/* Controles dentro de dialogos: cores legiveis garantidas
   (evita fundo == cor da letra em navegadores com modo escuro forcado). */
[data-testid="stDialog"] input, [data-testid="stDialog"] textarea,
[data-testid="stDialog"] div[data-baseweb="input"],
[data-testid="stDialog"] div[data-baseweb="input"] > div,
[data-testid="stDialog"] div[data-baseweb="select"] > div,
[data-testid="stDialog"] div[data-baseweb="select"] input,
[data-testid="stDialog"] [data-baseweb="select"] span,
[data-testid="stDialog"] [data-baseweb="select"] div {
    background: var(--card-bg) !important; color: var(--cor-texto) !important;
}
/* slider de cores (select_slider): o balao do valor selecionado fica sobre
   o fundo primario -> texto branco; o thumb e a trilha usam a cor primaria */
[data-testid="stDialog"] [data-baseweb="slider"] {
    color: var(--cor-texto) !important;
}
[data-testid="stDialog"] [data-baseweb="slider"] [role="slider"],
[data-testid="stDialog"] [data-baseweb="slider"] [role="slider"] * {
    background-color: var(--cor-p) !important;
    color: #ffffff !important;
}
[data-testid="stDialog"] [data-baseweb="slider"] [data-testid="stSliderThumbValue"],
[data-testid="stDialog"] [data-baseweb="slider"] [data-testid="stSliderThumbValue"] * {
    background-color: transparent !important;
    color: #ffffff !important;
}
[data-testid="stDialog"] [data-testid="stSliderThumbValue"] [data-testid="stMarkdownContainer"] {
    background-color: transparent !important;
}
[data-testid="stDialog"] [data-testid="stSliderTickBar"] [data-testid="stMarkdownContainer"] {
    color: var(--cor-texto) !important;
}
[data-testid="stDialog"] [data-baseweb="popover"] [role="listbox"],
[data-testid="stDialog"] [data-baseweb="menu"],
[data-testid="stDialog"] [data-baseweb="popover"] ul {
    background: var(--card-bg) !important; color: var(--cor-texto) !important;
}
[data-testid="stDialog"] [data-baseweb="popover"] li {
    color: var(--cor-texto) !important;
}
[data-testid="stDialog"] [data-baseweb="popover"] li:hover,
[data-testid="stDialog"] [data-baseweb="popover"] li[aria-selected="true"] {
    background: @@CORS_SOFT@@ !important; color: var(--cor-texto) !important;
}
[data-testid="stDialog"] .stButton > button[kind="secondary"] {
    background: var(--card-bg) !important; color: var(--cor-texto) !important;
    border: 1px solid var(--borda) !important;
}
.stAlert {border-radius: 10px; border: 1px solid var(--borda);}
[data-testid="stDataFrame"] {border: 1px solid var(--borda); border-radius: 12px; overflow: hidden;}

/* ---------------- Divisores ---------------- */
hr {margin: .7rem 0; border-color: var(--borda);}

/* =====================================================================
   RESPONSIVO
   Estrategia: em telas <= 720px as colunas que nao cabem quebram linha
   (flex-wrap), empilhando formularios e cartoes; o calendario do
   dashboard permanece com as 7 colunas de dias; a grade semanal usa 2
   colunas por linha; a tabela do calendario de planos rola na horizontal.
   ===================================================================== */
@media (max-width: 720px) {
    .block-container {padding: .6rem .9rem 2rem;}
    h1 {font-size: 1.45rem !important;}
    h2 {font-size: 1.3rem !important;}
    h3 {font-size: 1.15rem !important;}
    .card-titulo {font-size: .9rem;}
    div[data-testid="stMetricValue"] {font-size: 1.35rem !important;}
    div[data-testid="stMetric"] {padding: 10px 12px;}
    .stTabs [data-baseweb="tab"] {padding: 8px 10px; font-size: .85rem;}
    div[class*="st-key-card_"] {padding: 12px;}
    .gcell {padding: 6px 8px;}
    .postit {padding: 10px 12px;}

    div[data-testid="stHorizontalBlock"] {flex-wrap: wrap !important;}
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 260px !important; min-width: 0 !important; width: auto !important;
    }
    /* Calendario do Dashboard: mantem 7 colunas compactas */
    div[class*="st-key-card_dash_cal"] div[data-testid="stColumn"] {
        flex: 1 1 0 !important; min-width: 0 !important; width: 100% !important;
    }
    div[class*="st-key-cal_d_"] button {
        padding: .1rem 0 !important; min-height: 1.6rem; font-size: .62rem; border-radius: 8px;
        white-space: nowrap !important; overflow: hidden; font-variant-numeric: tabular-nums;
        letter-spacing: 0; overflow-wrap: normal !important; word-break: normal !important;
        text-overflow: clip;
    }
    div[class*="st-key-cal_d_"] button [data-testid="stMarkdownContainer"],
    div[class*="st-key-cal_d_"] button [data-testid="stMarkdownContainer"] p {
        white-space: nowrap !important;
        overflow-wrap: normal !important;
        word-break: normal !important;
        font-size: .62rem !important;
        margin: 0 !important;
    }
    /* Grade Semanal: 2 dias por linha em telas pequenas */
    div[class*="st-key-grade_dias"] div[data-testid="stColumn"] {
        flex: 1 1 45% !important; min-width: 45% !important; width: auto !important;
    }
    /* Abas Entrar/Criar do login lado a lado */
    div[class*="st-key-login_card"] div[data-testid="stColumn"] {
        flex: 1 1 0 !important; min-width: 0 !important; width: 100% !important;
    }
    .cal-table {border-spacing: 2px;}
    .cal-table th {font-size: .6rem; padding: 2px;}
    .cal-table td {padding: 5px 0 !important; font-size: .75rem; border-radius: 6px;}
    [data-testid="stDialog"] {max-width: 94vw !important;}
}

/* Botao flutuante do menu (criado pelo JS injetado).
   Visibilidade controlada pelo JS: no celular fica sempre visivel;
   no desktop aparece so quando a sidebar esta oculta (para poder reabri-la). */
#ei-mobile-menu {
    position: fixed; right: 14px; bottom: 14px; z-index: 99999;
    width: 52px; height: 52px; border-radius: 16px; border: none; cursor: pointer;
    background: linear-gradient(135deg, @@CORP@@, @@CORP_DEEP@@) !important;
    color: @@BTN@@ !important; font-size: 24px; line-height: 1;
    align-items: center; justify-content: center;
    box-shadow: 0 8px 22px @@CORP_SHADOW@@;
    -webkit-tap-highlight-color: transparent;
}
#ei-mobile-menu:active { transform: scale(.94); }

@media (max-width: 480px) {
    div[class*="st-key-login_card"] {padding: 1.3rem 1.1rem !important;}
    [data-testid="stMain"] .block-container {padding-top: 2vh !important;}
    body::before, body::after {opacity: .55;}
    .cal-scroll .cal-table {min-width: 460px;}
}
</style>
"""

def injetar_css(config):
    aparencia = config.get("aparencia", "System")
    dark = aparencia == "Dark"
    cor_tema = config.get("cor_tema", "blue")
    presets = {
        "blue": ("#1f538d", "#14375e"),
        "green": ("#1d7a46", "#104f2c"),
        "dark-blue": ("#1a3a6b", "#0e2340"),
    }
    if cor_tema in presets:
        cor_p, cor_s = presets[cor_tema]
    else:
        cor_p = config.get("cor_principal", "#1f538d")
        cor_s = config.get("cor_secundaria", "#14375e")
    cor_fundo = config.get("cor_fundo", "#2b2b2b") if dark else "#e4e7ec"
    txt_btn = cor_texto_legivel(cor_p)
    card_bg = "#2b2b2b" if dark else "#ffffff"
    borda = "rgba(255,255,255,0.10)" if dark else "#d3d8e0"
    texto = "#e8e8e8" if dark else "#1f1f1f"
    texto_cinza = "#9a9a9a" if dark else "#6c757d"

    cor_p_hover = ajustar_cor(cor_p, 1.12)
    cor_p_deep = ajustar_cor(cor_p, 0.82)
    cor_p_edge = ajustar_cor(cor_p, 0.66)
    cor_s_grad = ajustar_cor(cor_s, 0.80)
    tag_txt = ajustar_cor(cor_s, 0.55)
    cor_borda = (config.get("cor_borda_card") or "").strip() or cor_p

    css = CSS_TEMPLATE
    css = css.replace("@@SCHEME@@", "dark" if dark else "light")
    css = css.replace("@@CORP@@", cor_p)
    css = css.replace("@@CORS@@", cor_s)
    css = css.replace("@@CORP_HOVER@@", cor_p_hover)
    css = css.replace("@@CORP_DEEP@@", cor_p_deep)
    css = css.replace("@@CORP_EDGE@@", cor_p_edge)
    css = css.replace("@@CORP_SOFT@@", cor_rgba(cor_p, 0.18))
    css = css.replace("@@CORP_SHADOW@@", cor_rgba(cor_p, 0.35))
    css = css.replace("@@CORS_GRAD@@", cor_s_grad)
    css = css.replace("@@CORS_SOFT@@", cor_rgba(cor_s, 0.12))
    css = css.replace("@@COR_BORDA@@", cor_borda)
    css = css.replace("@@TAGS@@", tag_txt)
    css = css.replace("@@SCROLL@@", cor_rgba(cor_p, 0.45))
    css = css.replace("@@MINI_DIA@@", cor_rgba(cor_s, 0.06))
    css = css.replace("@@MINI_AULA@@", cor_rgba(cor_s, 0.14))
    css = css.replace("@@MINI_TXT@@", tag_txt)
    css = css.replace("@@FUNDO@@", cor_fundo)
    css = css.replace("@@BTN@@", txt_btn)
    css = css.replace("@@CARDBG@@", card_bg)
    css = css.replace("@@BORDA@@", borda)
    css = css.replace("@@TEXTO@@", texto)
    css = css.replace("@@CINZA@@", texto_cinza)
    st.markdown(css, unsafe_allow_html=True)

CSS_LOGIN = """
<style>
/* ---------- Tela de login: fundo ---------- */
html, body, .stApp, section[data-testid="stMain"] {
    background: linear-gradient(135deg, #edf1ff 0%, #e2e9fd 32%, #f3f0ff 66%, #fdeef6 100%) !important;
}
[data-testid="stSidebar"] { display: none !important; }
[data-testid="stMain"] .block-container {
    max-width: 470px !important;
    padding-top: 5vh !important;
    padding-bottom: 4vh !important;
}
/* bolhas decorativas de fundo */
body::before, body::after {
    content: ""; position: fixed; border-radius: 50%; filter: blur(70px);
    z-index: 0; pointer-events: none; opacity: .85;
}
body::before {
    width: 46vw; height: 46vw; min-width: 380px; min-height: 380px;
    top: -12vw; left: -10vw;
    background: radial-gradient(circle, rgba(79,70,229,.22), rgba(79,70,229,0) 70%);
}
body::after {
    width: 48vw; height: 48vw; min-width: 400px; min-height: 400px;
    bottom: -14vw; right: -12vw;
    background: radial-gradient(circle, rgba(14,165,233,.18), rgba(14,165,233,0) 70%);
}
/* ---------- Card central ---------- */
div[class*="st-key-login_card"] {
    background: rgba(255,255,255,.94) !important;
    border: 1px solid rgba(255,255,255,.75) !important;
    border-radius: 24px !important;
    padding: 2.1rem 2.2rem 1.9rem !important;
    box-shadow: 0 26px 60px rgba(31,83,141,.16), 0 4px 14px rgba(0,0,0,.05) !important;
    position: relative; z-index: 1;
}
/* ---------- Marca (legivel em fundo claro) ---------- */
.login-logo { text-align: center; margin: 0 auto 10px auto; line-height: 0; }
.login-img {
    width: 300px; max-width: 100%; height: auto; display: inline-block;
    filter: drop-shadow(0 10px 24px rgba(31,83,141,.25));
}
.login-titulo {
    text-align: center; font-size: 1.45rem; font-weight: 800; color: #111827 !important;
    line-height: 1.15; letter-spacing: -.02em; margin-bottom: 4px;
}
.login-titulo small {
    display: block; margin-top: 4px; font-size: .6rem; font-weight: 700;
    letter-spacing: 4px; text-transform: uppercase; color: #4f46e5 !important;
}
.login-sub { color: #6b7280 !important; font-size: .9rem; margin: 4px 0 1.5rem 0; }
.login-foot {
    margin-top: 1.6rem; text-align: center; color: #9ca3af !important;
    font-size: .78rem; border-top: 1px solid #f0f2f5; padding-top: 1.1rem;
}
/* ---------- Abas Entrar / Criar Conta ---------- */
div[class*="st-key-tab_entrar"] button, div[class*="st-key-tab_criar"] button {
    border-radius: 12px !important; font-weight: 700 !important;
    border: 1.5px solid #e2e8f0 !important; background: #ffffff !important;
    color: #64748b !important; box-shadow: none !important; height: 42px;
    transition: all .18s ease;
}
div[class*="st-key-tab_entrar"] button[kind="primary"],
div[class*="st-key-tab_criar"] button[kind="primary"] {
    background: linear-gradient(135deg, #1f538d, #4f46e5) !important;
    border: none !important; color: #ffffff !important;
    box-shadow: 0 8px 18px rgba(31,83,141,.28) !important;
}
/* ---------- Inputs ---------- */
[data-testid="stTextInput"] label p {
    color: #374151 !important; font-weight: 600; font-size: .83rem;
}
/* Campo unico: o contorno/contanier BaseWeb fica transparente e o <input>
   real assume a borda/fundo, evitando o "retangulo dentro do retangulo"
   (campo branco com area de digitacao cinza deslocada) que desalinha os
   icones de gerenciadores de senha. */
[data-testid="stTextInput"] div[data-baseweb="input"],
[data-testid="stTextInput"] div[data-baseweb="input"] > div {
    background: transparent !important;
    border: none !important;
    padding: 0 !important;
}
[data-testid="stTextInput"] input {
    background: #f8fafc !important; border: 1.5px solid #e2e8f0 !important;
    border-radius: 12px !important; padding: .7rem .9rem !important;
    color: #111827 !important; font-size: .95rem !important;
    width: 100% !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: #4f46e5 !important; background: #ffffff !important;
    box-shadow: 0 0 0 4px rgba(79,70,229,.12) !important;
    outline: none !important;
}
[data-testid="stTextInput"] input::placeholder { color: #9ca3af !important; opacity: 1; }
/* ---------- Botao CTA ---------- */
.stFormSubmitButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1f538d, #4f46e5) !important;
    border: none !important; border-radius: 12px !important;
    height: 48px !important; font-weight: 700 !important; font-size: .95rem !important;
    box-shadow: 0 12px 24px rgba(31,83,141,.30) !important;
    transition: transform .18s ease, box-shadow .18s ease;
}
.stFormSubmitButton > button[kind="primary"]:hover {
    transform: translateY(-1px); box-shadow: 0 16px 30px rgba(31,83,141,.38) !important;
}
/* ---------- Alertas ---------- */
[data-testid="stAlert"] { border-radius: 12px !important; font-size: .85rem; }
</style>
"""

def injetar_css_login():
    st.markdown(CSS_LOGIN, unsafe_allow_html=True)

def html_card(titulo, sub="", tag="", extra_html=""):
    partes = []
    partes.append('<div class="card">')
    if tag:
        partes.append(f'<span class="tag">{tag}</span>')
    partes.append(f'<div class="card-titulo">{titulo}</div>')
    if sub:
        partes.append(f'<div class="card-sub">{sub}</div>')
    if extra_html:
        partes.append(extra_html)
    partes.append('</div>')
    return "".join(partes)

# =====================================================================
# 3. MOTOR DE GERACAO DO WORD (em memoria, para download)
# =====================================================================
def gerar_documento_word(questoes_selecionadas, config, incluir_gabarito,
                         modo_acessibilidade, mostrar_descritor, titulo_prova):
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = config.get("fonte", "Arial")
    style.paragraph_format.space_after = Pt(config.get("espacamento_pt", 0))

    margem_inches = config.get("margem_cm", 1.5) / 2.54
    for section in doc.sections:
        section.top_margin = Inches(margem_inches)
        section.bottom_margin = Inches(margem_inches)
        section.left_margin = Inches(margem_inches)
        section.right_margin = Inches(margem_inches)

    if modo_acessibilidade:
        font.size = Pt(16)
    else:
        font.size = Pt(config.get("tamanho_fonte", 11))

        tabela_cab = doc.add_table(rows=3, cols=3)
        tabela_cab.style = 'Table Grid'

        celula_escola = tabela_cab.cell(0, 0)
        celula_escola.merge(tabela_cab.cell(0, 2))
        p_escola = celula_escola.paragraphs[0]
        p_escola.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_escola = p_escola.add_run(f"{config.get('escola', 'ESCOLA PADRAO')}")
        run_escola.bold = True
        run_escola.font.size = Pt(config.get("tamanho_fonte", 11) + 2)

        if config.get("mostrar_aluno", True):
            celula_aluno = tabela_cab.cell(1, 0)
            celula_aluno.merge(tabela_cab.cell(1, 2))
            p_aluno = celula_aluno.paragraphs[0]
            p_aluno.add_run(f" ALUNO(A): {'_' * 75}")

        celula_turma = tabela_cab.cell(2, 0)
        if config.get("mostrar_turma", True):
            celula_turma.paragraphs[0].add_run(" TURMA: ____________")

        celula_data = tabela_cab.cell(2, 1)
        if config.get("mostrar_data", True):
            celula_data.paragraphs[0].add_run(" DATA: ___/___/20___")

        celula_prof = tabela_cab.cell(2, 2)
        prof_nome = config.get('professor', '').strip()
        if prof_nome:
            celula_prof.paragraphs[0].add_run(f" PROF.: {prof_nome}")

        doc.add_paragraph()
        if titulo_prova.strip():
            p_titulo_prova = doc.add_paragraph()
            p_titulo_prova.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_tit = p_titulo_prova.add_run(titulo_prova.upper())
            run_tit.bold = True
            run_tit.font.size = Pt(config.get("tamanho_fonte", 11) + 2)

        doc.add_paragraph()

        if config.get("usar_duas_colunas", True):
            new_section = doc.add_section(WD_SECTION_START.CONTINUOUS)
            sectPr = new_section._sectPr
            cols = parse_xml(r'<w:cols w:num="2" w:space="708" w:sep="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            sectPr.append(cols)

    for i, q in enumerate(questoes_selecionadas):
        p_q = doc.add_paragraph()
        p_q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        tag_descritor = ""
        if mostrar_descritor and str(q.get('tema', '')).startswith("D") and " - " in q.get('tema', ''):
            codigo = q['tema'].split(" - ")[0]
            tag_descritor = f" ({codigo})"

        num_formatado = str(i + 1).zfill(2)
        p_q.add_run(f"{num_formatado}.{tag_descritor} ").bold = True

        if q.get('enunciado', '').strip():
            linhas_enunciado = q['enunciado'].strip().split('\n')
            p_q.add_run(linhas_enunciado[0])
            for linha in linhas_enunciado[1:]:
                p_extra = doc.add_paragraph(linha)
                p_extra.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        caminho_img = q.get('imagem', '')
        if caminho_img and imagem_existe(caminho_img):
            try:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img = p_img.add_run()

                dados_img = imagem_bytes(caminho_img)
                img_temp = Image.open(io.BytesIO(dados_img))
                largura_px, altura_px = img_temp.size
                img_temp.close()

                limite_largura = 7.0 if config.get("usar_duas_colunas", True) else 10.0
                limite_altura = 6.0
                proporcao = largura_px / altura_px if altura_px else 1

                if proporcao >= 1.2:
                    run_img.add_picture(io.BytesIO(dados_img), width=Cm(limite_largura))
                else:
                    run_img.add_picture(io.BytesIO(dados_img), height=Cm(limite_altura))
            except Exception:
                p_erro = doc.add_paragraph("[Aviso: Erro ao renderizar a imagem.]")
                p_erro.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if q.get('pergunta_direta', '').strip():
            p_pergunta = doc.add_paragraph()
            p_pergunta.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_pergunta.add_run(q['pergunta_direta']).bold = True

        alternativas = q.get('alternativas', {})
        if len(alternativas) > 0:
            alts_keys = list(alternativas.keys())
            if modo_acessibilidade:
                for letra in alts_keys:
                    p_alt = doc.add_paragraph(f"{letra}) {alternativas[letra]}")
                    p_alt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                for letra in alts_keys:
                    p_alt = doc.add_paragraph(f"{letra}) {alternativas[letra]}")
                    p_alt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p_alt.paragraph_format.left_indent = Inches(0.2)
            doc.add_paragraph()
        else:
            p_linhas = doc.add_paragraph("\n________________________________________________________")
            p_linhas.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_linhas2 = doc.add_paragraph("________________________________________________________\n")
            p_linhas2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if incluir_gabarito:
        doc.add_page_break()
        gab_section = doc.add_section(WD_SECTION_START.CONTINUOUS)
        gab_sectPr = gab_section._sectPr
        gab_cols = parse_xml(r'<w:cols w:num="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        gab_sectPr.append(gab_cols)

        p_gab = doc.add_paragraph()
        p_gab.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_gab.add_run("GABARITO OFICIAL\n").bold = True

        tabela_gab = doc.add_table(rows=1, cols=2)
        tabela_gab.style = 'Table Grid'
        hdr_cells = tabela_gab.rows[0].cells
        hdr_cells[0].text = 'Questao'
        hdr_cells[1].text = 'Resposta'

        for i, q in enumerate(questoes_selecionadas):
            num_formatado = str(i + 1).zfill(2)
            row_cells = tabela_gab.add_row().cells
            row_cells[0].text = f"Questao {num_formatado}"
            row_cells[1].text = q.get('gabarito', '')

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# 3.1 JS INJETADO (componente): menu movel + swipe + cookie de login
#     Roda dentro de um iframe do mesmo dominio e controla o pai.
# =====================================================================
def _html_js_movel(usuario=None, token=None):
    import json as _json

    payload = None
    if usuario and token:
        payload = {"usuario": usuario, "token": token}

    _BODY = r"""(function(){
  var d = document;
  var W = window;
  var COOKIE = "ei_usuario";

  function isMobile() { return matchMedia("(max-width: 768px)").matches; }
  function el(sel) { return d.querySelector(sel); }
  function side() { return el('[data-testid="stSidebar"]'); }

  // Aberta = transform 'none' (o botao de recolher permanece no DOM mesmo recolhido).
  function isOpen() {
    var s = side();
    if (!s) return false;
    var t = getComputedStyle(s).transform || '';
    if (t === 'none' || t === '') return true;
    var m = t.match(/matrix\(([-\d.,\s]+)\)/);
    if (!m) return true;
    var tx = parseFloat(m[1].split(',')[4]) || 0;
    return tx > -1;
  }
  // Streamlit expoe um <div> wrapper; o clique efetivo e no <button> interno.
  function sbBtn(t) { var w = el('[data-testid="' + t + '"]'); if (!w) return null; return w.querySelector('button') || w; }
  function openSb() { var b = sbBtn('stExpandSidebarButton'); if (b) b.click(); }
  function closeSb() { var b = sbBtn('stSidebarCollapseButton'); if (b) b.click(); }
  function toggleSb() { isOpen() ? closeSb() : openSb(); }

  // --- menu flutuante (sempre que a sidebar estiver oculta; no celular, sempre) ---
  function ensureFab() {
    var side = el('[data-testid="stSidebar"]');
    if (!side || !side.querySelector('button')) return;
    if (el('#ei-mobile-menu')) return;
    var fab = d.createElement('button');
    fab.id = 'ei-mobile-menu';
    fab.innerHTML = '&#9776;';
    fab.setAttribute('aria-label', 'Abrir menu');
    fab.addEventListener('click', function(ev) { ev.stopPropagation(); ev.preventDefault(); toggleSb(); });
    d.body.appendChild(fab);
  }
  function syncFab() {
    var fab = el('#ei-mobile-menu');
    if (!fab) return;
    var show = isMobile() || !isOpen();
    fab.style.display = show ? 'flex' : 'none';
  }

  // --- gesto: arrastar a partir de qualquer lugar abre; empurrar p/ esq fecha ---
  var drag = null;

  // Bloqueia a navegacao de voltar por gesto (overscroll + touch-action),
  // no desktop impede recolher a barra e desenha a "alca" de arrastar no celular.
  function injectStyles() {
    if (d.getElementById('ei-styles')) return;
    var s = d.createElement('style');
    s.id = 'ei-styles';
    s.textContent =
      'html, body { overscroll-behavior-x: contain !important; }' +
      '@media (min-width: 769px) {' +
      ' [data-testid="stSidebarCollapseButton"], [data-testid="stExpandSidebarButton"] { display: none !important; }' +
      ' #ei-edge, #ei-mobile-menu { display: none !important; }' +
      ' section[data-testid="stSidebar"] { transform: none !important;' +
      ' width: 300px !important; min-width: 300px !important; max-width: 300px !important; }' +
      '}' +
      '#ei-edge { position: fixed; left: 0; top: 0; bottom: 0; width: 30px;' +
      ' z-index: 2147483646; pointer-events: none;' +
      ' display: flex; align-items: center; justify-content: center;' +
      ' -webkit-tap-highlight-color: transparent; }' +
      '#ei-edge::before { content: ""; width: 6px; height: 84px; border-radius: 6px;' +
      ' background: rgba(120,120,120,0.45); box-shadow: 0 0 0 5px rgba(120,120,120,0.10);' +
      ' animation: ei-edge-pulse 1.9s ease-in-out infinite; }' +
      '@keyframes ei-edge-pulse { 0%, 100% { opacity: .5; transform: translateX(0); }' +
      ' 50% { opacity: 1; transform: translateX(2px); } }';
    (d.head || d.documentElement).appendChild(s);
  }

  function ensureOverlay() {
    if (!isMobile() || el('#ei-edge')) return;
    if (!side()) return;
    var o = d.createElement('div');
    o.id = 'ei-edge';
    d.body.appendChild(o);
    syncOverlay();
  }
  function syncOverlay() {
    var o = el('#ei-edge');
    if (o) o.style.display = (isMobile() && !isOpen()) ? 'flex' : 'none';
    syncFab();
  }

  // largura real da sidebar (le o translateX do estado recolhido do Streamlit)
  function sbWidth() {
    var s = side();
    if (!s) return 300;
    var t = getComputedStyle(s).transform || '';
    var m = t.match(/matrix\(([-\d.,\s]+)\)/);
    if (m) {
      var tx = Math.abs(parseFloat(m[1].split(',')[4]) || 0);
      if (tx > 5) return tx;
    }
    return s.offsetWidth || 300;
  }

  function inXScroll(e) {
    if (!e.target || !e.target.closest) return false;
    return !!e.target.closest('.cal-scroll, [data-testid="stDataFrame"], [data-testid="stTable"]');
  }

  // Abrir do gesto em qualquer lugar, exceto dentro de areas de rolagem
  // horizontal que ainda tenham conteudo para tras (o gesto rola o conteudo).
  function canOpenHere(e) {
    if (!e.target || !e.target.closest) return true;
    var sc = e.target.closest('.cal-scroll, [data-testid="stDataFrame"], [data-testid="stTable"]');
    if (!sc) return true;
    return sc.scrollLeft <= 2;
  }

  function onStart(e) {
    if (!isMobile() || e.touches.length !== 1) return;
    var t = e.touches[0];
    drag = { sx: t.clientX, sy: t.clientY, dx: 0, dy: 0, mode: null, w: sbWidth() };
  }
  function onMove(e) {
    if (!drag) return;
    var t = e.touches[0];
    var dx = t.clientX - drag.sx;
    var dy = t.clientY - drag.sy;
    drag.dx = dx; drag.dy = dy;
    var open = isOpen();
    if (!drag.mode) {
      var horiz = Math.abs(dx) > 10 && Math.abs(dx) > Math.abs(dy) * 1.15;
      if (!horiz) return;
      if (!open && canOpenHere(e)) drag.mode = 'open';
      else if (open && !inXScroll(e)) drag.mode = 'close';
      else { drag = null; return; }
    }
    e.preventDefault();
    var s = side(); if (!s) return;
    var w = drag.w;
    s.style.transition = 'none';
    if (drag.mode === 'open') {
      // largura fixa (300px): o conteudo nao se adapta, apenas a tela vai
      // revelando a barra conforme o dedo puxa para a direita
      s.style.width = w + 'px'; s.style.minWidth = w + 'px'; s.style.maxWidth = w + 'px';
      s.style.transform = 'translateX(' + Math.max(-w, -w + drag.dx) + 'px)';
    } else {
      s.style.transform = 'translateX(' + Math.max(-w, drag.dx) + 'px)';
    }
  }
  function onEnd(e) {
    if (!drag) return;
    var was = drag; drag = null;
    var s = side(); if (!s) return;
    var w = was.w;
    var lim = Math.max(60, w * 0.25);
    if (was.mode === 'open') {
      if (was.dx > lim) commitOpen(s, w); else reset(s);
    } else if (was.mode === 'close') {
      if (was.dx < -lim) commitClose(s, w); else reset(s);
    }
    syncOverlay();
  }
  function commitOpen(s, w) {
    s.style.transition = 'none';
    s.style.width = w + 'px'; s.style.minWidth = w + 'px'; s.style.maxWidth = w + 'px';
    s.style.transform = 'translateX(0)';
    openSb();
    setTimeout(function() { reset(s); syncOverlay(); }, 400);
  }
  function commitClose(s, w) {
    s.style.transition = 'none';
    s.style.transform = 'translateX(-' + w + 'px)';
    closeSb();
    setTimeout(function() { reset(s); syncOverlay(); }, 400);
  }
  function reset(s) {
    s.style.transition = ''; s.style.width = ''; s.style.minWidth = ''; s.style.maxWidth = ''; s.style.transform = '';
  }

  d.addEventListener('touchstart', onStart, {passive: true});
  d.addEventListener('touchmove', onMove, {passive: false});
  d.addEventListener('touchend', onEnd, {passive: true});
  d.addEventListener('touchcancel', onEnd, {passive: true});

  // --- auto-ocultar apos navegar pelo menu (so no celular) ---
  function watchNav() {
    if (!isMobile()) return;
    var side = el('[data-testid="stSidebar"]');
    if (!side || side.__eiWatch) return;
    side.__eiWatch = true;
    side.addEventListener('click', function(ev) {
      var b = ev.target.closest('button');
      if (!b) return;
      if (b.getAttribute('data-testid') === 'stSidebarCollapseButton') return;
      setTimeout(function() { if (isMobile() && isOpen()) closeSb(); }, 450);
    });
  }

  // --- cookie de login (lembrar da conta) ---
  function setCookie(name, value) {
    d.cookie = name + "=" + encodeURIComponent(value) + "; path=/; max-age=2592000; SameSite=Lax";
  }
  function clearCookie(name) {
    d.cookie = name + "=; path=/; max-age=0";
  }
  function getCookie(name) {
    var m = d.cookie.match(new RegExp("(?:^|; )" + name + "=([^;]*)"));
    return m ? decodeURIComponent(m[1]) : "";
  }
  function syncCookie() {
    var p = window.__EI_PAYLOAD;
    if (p && p.usuario) { setCookie(COOKIE, p.usuario + "|" + p.token); }
    else { if (getCookie(COOKIE)) clearCookie(COOKIE); }
  }
  W.__eiSyncCookie = syncCookie;

  injectStyles();
  ensureFab();
  syncFab();
  watchNav();
  syncCookie();

  // No celular a barra abre sozinha assim que o app carrega (somente 1a vez).
  var didInitOpen = false;
  function forceInitialOpen() {
    if (didInitOpen || !isMobile()) return;
    var s = side();
    if (!s) return;
    if (isOpen()) { didInitOpen = true; return; }
    var b = sbBtn('stExpandSidebarButton');
    if (b) { b.click(); didInitOpen = true; }
  }
  // No desktop a barra e sempre visivel: se por qualquer motivo fechou, reabre.
  function enforceDesktop() {
    if (isMobile()) return;
    if (!side()) return;
    if (!isOpen()) { var b = sbBtn('stExpandSidebarButton'); if (b) b.click(); }
  }
  W.addEventListener('resize', function() { forceInitialOpen(); enforceDesktop(); watchNav(); syncOverlay(); });

  W.__ei = {
    isOpen: isOpen, openSb: openSb, closeSb: closeSb, toggleSb: toggleSb,
    sbWidth: sbWidth, isMobile: isMobile, syncCookie: syncCookie
  };

  var obs = new MutationObserver(function() {
    ensureFab(); watchNav(); ensureOverlay(); forceInitialOpen(); enforceDesktop(); syncOverlay();
  });
  if (d.body) obs.observe(d.body, { childList: true, subtree: true });
  setTimeout(function() { ensureOverlay(); forceInitialOpen(); enforceDesktop(); syncOverlay(); }, 400);
  setInterval(function() { ensureOverlay(); forceInitialOpen(); enforceDesktop(); syncOverlay(); }, 500);
})();
"""

    return f"""<script>
(function(){{
  var W = window.parent;
  var d = W.document;
  if (!d || !d.documentElement) return;
  try {{ W.__EI_PAYLOAD = {_json.dumps(payload)}; if (W.__eiSyncCookie) W.__eiSyncCookie(); }} catch (e) {{}}
  if (d.getElementById('ei-inj')) return;
  var s = d.createElement('script');
  s.id = 'ei-inj';
  s.textContent = 'window.__EI_PAYLOAD = (' + {_json.dumps(payload)} + ');\\n' + {_json.dumps(_BODY)};
  (d.head || d.documentElement).appendChild(s);
}})();
</script>"""


def injetar_js_movel(usuario=None, token=None):
    st.iframe(_html_js_movel(usuario, token), width=1, height=1)

# =====================================================================
# 4. MOTOR EXCEL (relatorio de avaliacao, em memoria)
# =====================================================================
def gerar_excel_avaliacao(av):
    gabarito = av['gabarito']
    qtd_q = len(gabarito)
    notas = av['notas_alunos']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatorio de Notas"

    font_bold = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    fill_adequado = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    fill_intermediario = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    fill_critico = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_muito_critico = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_vazio = PatternFill(start_color="00CCFF", end_color="00CCFF", fill_type="solid")

    col_fim = 5 + qtd_q

    def apply_border(cell):
        cell.border = border_thin

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_fim)
    c_titulo = ws.cell(row=1, column=1, value=f"ESCOLA DE ENSINO FUNDAMENTAL - {av['titulo'].upper()}")
    c_titulo.font = Font(bold=True, size=14)
    c_titulo.alignment = align_center

    ws.merge_cells("C3:E3")
    c = ws.cell(row=3, column=3, value="TURMA:")
    c.font = font_bold; c.alignment = align_center; apply_border(c)
    ws.merge_cells("F3:I3")
    c = ws.cell(row=3, column=6, value=av['turma'])
    c.alignment = align_center; apply_border(c)

    ws.merge_cells("K3:M3")
    c = ws.cell(row=3, column=11, value="DISCIPLINA")
    c.font = font_bold; c.alignment = align_center; apply_border(c)
    ws.merge_cells("N3:P3")
    c = ws.cell(row=3, column=14, value="MATEMATICA")
    c.alignment = align_center; apply_border(c)

    ws.merge_cells("S3:U3")
    c = ws.cell(row=3, column=19, value="DATA:")
    c.font = font_bold; c.alignment = align_center; apply_border(c)
    ws.merge_cells("V3:X3")
    c = ws.cell(row=3, column=22, value=av.get('data', ''))
    c.alignment = align_center; apply_border(c)

    for range_str in ["C3:E3", "F3:I3", "K3:M3", "N3:P3", "S3:U3", "V3:X3"]:
        for row in ws[range_str]:
            for cell in row:
                apply_border(cell)

    ws.merge_cells("F5:G5")
    c = ws.cell(row=5, column=6, value="SITUACAO")
    c.font = font_bold; c.alignment = align_center; apply_border(c)
    ws.cell(row=5, column=7).border = border_thin

    niveis_resumo = [("ADEQUADO", fill_adequado), ("INTERMEDIARIO", fill_intermediario),
                     ("CRITICO", fill_critico), ("MUITO CRITICO", fill_muito_critico)]
    for i, (nome_n, cor_n) in enumerate(niveis_resumo):
        r = 6 + i
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        c_n = ws.cell(row=r, column=6, value=nome_n)
        c_n.font = font_bold; c_n.fill = cor_n; c_n.alignment = align_center
        ws.cell(row=r, column=7).border = border_thin
        ws.cell(row=r, column=6).border = border_thin
        c_val = ws.cell(row=r, column=8, value=0)
        c_val.font = font_bold; c_val.alignment = align_center; apply_border(c_val)

    ws.merge_cells("K5:O5")
    c = ws.cell(row=5, column=11, value="Classifique as categorias")
    c.fill = fill_intermediario; c.alignment = align_center
    for cell in ws["K5:O5"][0]:
        apply_border(cell)

    legendas = [("ADEQUADO", fill_adequado, 90, 100), ("INTERMEDIARIO", fill_intermediario, 70, 90),
                ("CRITICO", fill_critico, 50, 70), ("MUITO CRITICO", fill_muito_critico, 0, 50)]
    for i, (nome_l, cor_l, v_min, v_max) in enumerate(legendas):
        r = 6 + i
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
        c_nome = ws.cell(row=r, column=11, value=nome_l)
        c_nome.font = font_bold; c_nome.alignment = align_center
        ws.cell(row=r, column=11).border = border_thin
        ws.cell(row=r, column=12).border = border_thin
        c_cor = ws.cell(row=r, column=13)
        c_cor.fill = cor_l; apply_border(c_cor)
        c_min = ws.cell(row=r, column=14, value=v_min)
        c_min.alignment = align_center; apply_border(c_min)
        c_a = ws.cell(row=r, column=15, value="A")
        c_a.alignment = align_center; apply_border(c_a)
        c_max = ws.cell(row=r, column=16, value=v_max)
        c_max.alignment = align_center; apply_border(c_max)

    r_cfg = 11
    ws.merge_cells(start_row=r_cfg, start_column=1, end_row=r_cfg, end_column=2)
    c = ws.cell(row=r_cfg, column=1, value="NUMERO DE QUESTOES")
    c.font = font_bold; apply_border(c)
    ws.cell(row=r_cfg, column=2).border = border_thin
    c_q = ws.cell(row=r_cfg, column=3, value=qtd_q)
    c_q.font = font_bold; c_q.alignment = align_center; apply_border(c_q)

    ws.merge_cells(start_row=r_cfg + 1, start_column=1, end_row=r_cfg + 1, end_column=2)
    c = ws.cell(row=r_cfg + 1, column=1, value="GABARITO")
    apply_border(c)
    ws.cell(row=r_cfg + 1, column=2).border = border_thin
    for i, q in enumerate(gabarito):
        c_gab = ws.cell(row=r_cfg + 1, column=3 + i, value=q['resposta_correta'].upper())
        c_gab.font = font_bold; c_gab.alignment = align_center; apply_border(c_gab)

    ws.merge_cells(start_row=r_cfg + 2, start_column=1, end_row=r_cfg + 2, end_column=2)
    c = ws.cell(row=r_cfg + 2, column=1, value="DESCRITOR")
    apply_border(c)
    ws.cell(row=r_cfg + 2, column=2).border = border_thin
    for i, q in enumerate(gabarito):
        c_desc = ws.cell(row=r_cfg + 2, column=3 + i, value=q['descritor'])
        c_desc.font = Font(size=9, bold=True)
        c_desc.alignment = align_center; apply_border(c_desc)

    r_hdr = 14
    headers = ["N", "NOME DO ALUNO"] + [str(i + 1) for i in range(qtd_q)] + ["ACT", "PORC", "SITUACAO"]
    for i, h in enumerate(headers):
        c_h = ws.cell(row=r_hdr, column=i + 1, value=h)
        c_h.font = font_bold
        c_h.alignment = align_center if i != 1 else align_left
        apply_border(c_h)

    alunos_ordenados = sorted(notas.items(), key=lambda x: x[0])
    distribuicao = {"ADEQUADO": 0, "INTERMEDIARIO": 0, "CRITICO": 0, "MUITO CRITICO": 0}

    r_aluno = 15
    for idx_a, (aluno, dados) in enumerate(alunos_ordenados, 1):
        c_n = ws.cell(row=r_aluno, column=1, value=idx_a)
        c_n.alignment = align_center; apply_border(c_n)
        c_nome = ws.cell(row=r_aluno, column=2, value=aluno)
        apply_border(c_nome)

        respostas_dadas = dados.get('respostas_dadas', '')
        resp_padded = respostas_dadas.ljust(qtd_q, ' ')
        acertos = 0

        for i, q in enumerate(gabarito):
            letra_marcada = resp_padded[i].upper()
            correta = q['resposta_correta'].upper()

            c_resp = ws.cell(row=r_aluno, column=3 + i)
            c_resp.alignment = align_center; apply_border(c_resp)

            if letra_marcada == ' ':
                c_resp.fill = fill_vazio
                c_resp.value = ""
            else:
                c_resp.value = letra_marcada
                if letra_marcada == correta:
                    c_resp.fill = fill_adequado
                    acertos += 1
                else:
                    c_resp.fill = fill_muito_critico

        c_act = ws.cell(row=r_aluno, column=3 + qtd_q, value=acertos)
        c_act.alignment = align_center; apply_border(c_act)

        perc = (acertos / qtd_q) * 100 if qtd_q > 0 else 0
        c_porc = ws.cell(row=r_aluno, column=4 + qtd_q, value=f"{perc:.1f}".replace('.', ','))
        c_porc.alignment = align_center; apply_border(c_porc)

        nivel = "MUITO CRITICO"
        if perc >= 90:
            nivel = "ADEQUADO"
        elif perc >= 70:
            nivel = "INTERMEDIARIO"
        elif perc >= 50:
            nivel = "CRITICO"

        distribuicao[nivel] += 1
        c_sit = ws.cell(row=r_aluno, column=5 + qtd_q, value=nivel)
        c_sit.alignment = align_center; apply_border(c_sit)
        r_aluno += 1

    ws.cell(row=6, column=8, value=distribuicao["ADEQUADO"])
    ws.cell(row=7, column=8, value=distribuicao["INTERMEDIARIO"])
    ws.cell(row=8, column=8, value=distribuicao["CRITICO"])
    ws.cell(row=9, column=8, value=distribuicao["MUITO CRITICO"])

    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 38
    for i in range(qtd_q):
        ws.column_dimensions[get_column_letter(3 + i)].width = 4
    ws.column_dimensions[get_column_letter(3 + qtd_q)].width = 6
    ws.column_dimensions[get_column_letter(4 + qtd_q)].width = 7
    ws.column_dimensions[get_column_letter(5 + qtd_q)].width = 18

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# 5. AJUDA DE NAVEGACAO (query params e links)
# =====================================================================
def ir_para(pagina):
    st.query_params["pagina"] = pagina

def pagina_atual():
    params = st.query_params.get_all("pagina")
    return params[0] if params else "Dashboard"

def val_param(nome, padrao=""):
    vals = st.query_params.get_all(nome)
    return vals[0] if vals else padrao

def limpar_params():
    st.query_params.clear()

def link_para(pagina, base="."):
    qs = []
    for k in list(st.query_params.keys()):
        for v in st.query_params.get_all(k):
            qs.append(f"{k}={v}")
    qs.append(f"pagina={pagina}")
    return base + "?" + "&".join(qs)

# =====================================================================
# 6. SIDEBAR (menu lateral)
# =====================================================================
NAV_PLANEJAMENTO = [
    ("Dashboard", "In\u00edcio"),
    ("Grade Semanal", "Grade Semanal"),
    ("Turmas & Alunos", "Turmas & Alunos"),
    ("Central de Planos", "Central de Planos"),
    ("Anotacoes", "Lembretes"),
]
NAV_AVALIACOES = [
    ("Notas e Estatisticas", "Notas e Estatisticas"),
    ("Cadastrar Questao", "Cadastrar Questao"),
    ("Importar IA", "Importar IA"),
    ("Catalogo de Questoes", "Catalogo de Questoes"),
    ("Gerar Prova", "Gerar Prova"),
    ("Configuracoes", "Configuracoes"),
]

ICONES_NAV = {
    "Dashboard": "\u2302",
    "Grade Semanal": "\u25a6",
    "Turmas & Alunos": "\u25c8",
    "Central de Planos": "\u2691\uFE0E",
    "Anotacoes": "\u270e\uFE0E",
    "Notas e Estatisticas": "\u25a5",
    "Cadastrar Questao": "\uff0b",
    "Importar IA": "\u2913",
    "Catalogo de Questoes": "\u2630",
    "Gerar Prova": "\u25a4",
    "Configuracoes": "\u2699\uFE0E",
}

def montar_sidebar():
    atual = pagina_atual()
    with st.sidebar:
        st.markdown(
            '<div class="logo-aba">'
            f'<img src="data:image/png;base64,{LOGO_ABA_B64}" alt="Exame Inteligente" /></div>',
            unsafe_allow_html=True)
        st.markdown('<div class="nav-secao">Planejamento</div>', unsafe_allow_html=True)
        for chave, rotulo in NAV_PLANEJAMENTO:
            icone = ICONES_NAV.get(chave, "")
            rotulo_html = f"{icone}  {rotulo}" if icone else rotulo
            if st.button(rotulo_html, key=f"nav_p_{chave}",
                         type="primary" if atual == chave else "secondary",
                         use_container_width=True):
                st.query_params["pagina"] = chave
                st.rerun()
        st.markdown('<div class="nav-secao">Avaliacoes</div>', unsafe_allow_html=True)
        for chave, rotulo in NAV_AVALIACOES:
            icone = ICONES_NAV.get(chave, "")
            rotulo_html = f"{icone}  {rotulo}" if icone else rotulo
            if st.button(rotulo_html, key=f"nav_a_{chave}",
                         type="primary" if atual == chave else "secondary",
                         use_container_width=True):
                st.query_params["pagina"] = chave
                st.rerun()
        st.markdown("---")
        config = carregar_config()
        escola = config.get("escola", "Escola")
        st.caption(f"{escola}\n\nExame Inteligente 4.0 Web")
        user = usuario_atual()
        if user:
            conta = conta_atual(user)
            nome_conta = conta.get("nome", user) if conta else user
            st.caption(f"Logado: {nome_conta}\n@{user}")
            if st.button("Sair", key="sair_conta", use_container_width=True):
                st.session_state.pop("usuario", None)
                st.session_state.pop("bem_vindo", None)
                st.session_state.pop("login_modo", None)
                st.session_state["deslogado_manual"] = True
                st.query_params.clear()
                st.rerun()

# =====================================================================
# 6.1 TELA DE LOGIN / CRIACAO DE CONTA
# =====================================================================
def tela_login():
    injetar_css_login()
    with st.container(key="login_card"):
        st.markdown(
            '<div class="login-logo">'
            f'<img class="login-img" src="data:image/png;base64,{LOGO_INICIO_B64}" '
            'alt="Exame Inteligente" /></div>'
            '<div class="login-sub">Entre com a sua conta para acessar turmas, '
            'planos de aula, questoes e provas.</div>',
            unsafe_allow_html=True)

        modo = st.session_state.get("login_modo", "Entrar")
        c_tab = st.columns(2, gap="small")
        if c_tab[0].button("Entrar", key="tab_entrar",
                           type="primary" if modo == "Entrar" else "secondary",
                           use_container_width=True):
            st.session_state["login_modo"] = "Entrar"
            st.rerun()
        if c_tab[1].button("Criar Conta", key="tab_criar",
                           type="primary" if modo == "Criar Conta" else "secondary",
                           use_container_width=True):
            st.session_state["login_modo"] = "Criar Conta"
            st.rerun()

        if modo == "Entrar":
            with st.form("form_login"):
                login_user = st.text_input("Usuario", key="login_user",
                                           placeholder="seu.usuario")
                login_senha = st.text_input("Senha", type="password", key="login_senha",
                                            placeholder="digite sua senha")
                entrar = st.form_submit_button("Entrar na conta", type="primary",
                                               use_container_width=True)
            if entrar:
                conta = autenticar_usuario(login_user, login_senha)
                if conta:
                    st.session_state["usuario"] = conta["usuario"]
                    st.query_params.clear()
                    st.rerun()
                else:
                    st.error("Usuario ou senha incorretos.")
        else:
            with st.form("form_cadastro"):
                cad_user = st.text_input("Nome de usuario", key="cad_user",
                                         placeholder="ex.: maria.silva")
                cad_senha = st.text_input("Crie uma senha", type="password",
                                          key="cad_senha", placeholder="Minimo 6 caracteres")
                cad_confirmar = st.text_input("Confirme a senha", type="password",
                                              key="cad_confirmar", placeholder="Digite novamente")
                criar = st.form_submit_button("Criar minha conta", type="primary",
                                              use_container_width=True)
            if criar:
                user_c = cad_user.strip()
                if not user_c:
                    st.error("Preencha o nome de usuario.")
                elif len(cad_senha) < 6:
                    st.error("A senha precisa ter pelo menos 6 caracteres.")
                elif cad_senha != cad_confirmar:
                    st.error("As senhas nao conferem.")
                elif usuario_existe(user_c):
                    st.error("Ja existe uma conta com esse nome de usuario.")
                else:
                    criar_conta("", user_c, cad_senha)
                    st.session_state["usuario"] = user_c
                    st.session_state["bem_vindo"] = True
                    st.query_params.clear()
                    st.rerun()

        st.markdown(
            '<div class="login-foot">Exame Inteligente 4.0 &middot; '
            'Seus dados ficam guardados na sua conta.</div>',
            unsafe_allow_html=True)

# =====================================================================
# 7. CALENDARIO EM HTML
# =====================================================================
def parse_cal(cal):
    hoje = datetime.now()
    try:
        mes, ano = [int(x) for x in cal.split("/")]
    except Exception:
        mes, ano = hoje.month, hoje.year
    if mes < 1 or mes > 12:
        mes, ano = hoje.month, hoje.year
    return mes, ano

def html_grade_mini(grade_horaria):
    if not grade_horaria:
        return '<div class="card-sub">Grade vazia. Cadastre aulas na Grade Semanal.</div>'
    ordem = {d: i for i, d in enumerate(DIAS_UTEIS + ["Sabado", "Domingo"])}
    dias = sorted(list(set([i["dia"] for i in grade_horaria])), key=lambda x: ordem.get(x, 99))
    partes = ['<div class="gmini">']
    for dia in dias:
        aulas = [i for i in grade_horaria if i["dia"] == dia]
        aulas.sort(key=lambda x: x.get("aula", ""))
        cards = "".join(
            f'<div class="gmini-aula">{a.get("aula","")}-{a.get("turma","")}</div>'
            for a in aulas)
        partes.append(
            f'<div class="gmini-dia">'
            f'<div class="gmini-nome">{dia.split("-")[0]}</div>{cards}</div>')
    partes.append("</div>")
    return "".join(partes)

# =====================================================================
# 8. DIALOGOS REUTILIZAVEIS
# =====================================================================
CORES_POSTIT = [
    ("Amarelo Classico", "#fff3a3"),
    ("Verde Menta", "#b5e7a0"),
    ("Azul Ceu", "#a0c4ff"),
    ("Lilas Suave", "#cdb4db"),
    ("Rosa Pastel", "#ffc6ff"),
]

@st.dialog("Novo / Editar Post-it")
def dialog_postit(nota=None):
    grade = carregar_grade()
    turmas = ["Geral"] + sorted(list(set([i["turma"] for i in grade])))
    is_edicao = nota is not None

    titulo = st.text_input("Titulo do lembrete:", value=nota.get("titulo", "") if is_edicao else "")
    turma = st.selectbox("Atrelar a turma:", turmas,
                         index=turmas.index(nota["turma"]) if is_edicao and nota.get("turma") in turmas else 0)
    nomes_cores = [n for n, _ in CORES_POSTIT]
    cores_vals = [c for _, c in CORES_POSTIT]
    cor_padrao = nota.get("cor", "#fff3a3") if is_edicao else "#fff3a3"
    idx_cor = cores_vals.index(cor_padrao) if cor_padrao in cores_vals else 0
    cor = st.select_slider("Cor do post-it:", options=nomes_cores, value=nomes_cores[idx_cor])
    cor_hex = cores_vals[nomes_cores.index(cor)]
    conteudo = st.text_area("Anotacao / Lembrete:",
                            value=nota.get("conteudo", "") if is_edicao else "")

    if st.button("Salvar Post-it", type="primary", use_container_width=True):
        if not titulo.strip() or not conteudo.strip():
            st.error("Preencha o titulo e o conteudo.")
            return
        anotacoes = carregar_anotacoes()
        if is_edicao:
            for n in anotacoes:
                if n.get("id") == nota.get("id"):
                    n["titulo"] = titulo.strip()
                    n["conteudo"] = conteudo.strip()
                    n["turma"] = turma
                    n["cor"] = cor_hex
        else:
            novo_id = max([n.get("id", 0) for n in anotacoes], default=0) + 1
            anotacoes.append({
                "id": novo_id, "titulo": titulo.strip(), "conteudo": conteudo.strip(),
                "turma": turma, "cor": cor_hex,
                "data_criacao": datetime.now().strftime("%d/%m/%Y")
            })
        salvar_anotacoes(anotacoes)
        st.rerun()

@st.dialog("Visualizar / Editar Plano")
def dialog_plano(data_str, index_plano):
    planos = carregar_planos()
    if data_str not in planos or index_plano >= len(planos[data_str]):
        st.info("Plano nao encontrado.")
        return
    plano = dict(planos[data_str][index_plano])

    st.markdown(
        f"**Data:** {data_str} | **Horario:** {plano.get('horario','')} | "
        f"**Turma:** {plano.get('turma','')} | **Disciplina:** {plano.get('disciplina','Geral')}")

    novo_tema = st.text_input("Titulo / Tema da aula:", value=plano.get("tema", ""))

    campos = [
        ("Metodologia / Atividades", "metodologia"),
        ("Objetivos Especificos", "objetivos"),
        ("Procedimentos", "procedimentos"),
        ("Habilidade(s) da BNCC", "habilidades"),
        ("Competencia Geral", "comp_geral"),
        ("Competencias Especificas", "comp_especifica"),
        ("Recursos Necessarios", "recursos"),
        ("Avaliacao", "avaliacao"),
        ("Observacoes", "observacoes"),
    ]
    novos_valores = {}
    abas = st.tabs(["Metodologia & Objetivos", "Estrutura BNCC", "Recursos & Avaliacao"])
    mapa_abas = {
        "metodologia": 0, "objetivos": 0, "procedimentos": 0,
        "habilidades": 1, "comp_geral": 1, "comp_especifica": 1,
        "recursos": 2, "avaliacao": 2, "observacoes": 2,
    }
    for rotulo, chave in campos:
        with abas[mapa_abas[chave]]:
            novos_valores[chave] = st.text_area(rotulo, value=plano.get(chave, ""))

    c1, c2 = st.columns(2)
    if c1.button("Salvar Alteracoes", type="primary", use_container_width=True):
        if not novo_tema.strip():
            st.error("O tema da aula e obrigatorio.")
            return
        plano["tema"] = novo_tema.strip()
        for chave, valor in novos_valores.items():
            plano[chave] = valor.strip()
        planos[data_str][index_plano] = plano
        salvar_planos(planos)
        st.rerun()
    if c2.button("Excluir Plano", use_container_width=True):
        del planos[data_str][index_plano]
        if len(planos[data_str]) == 0:
            del planos[data_str]
        salvar_planos(planos)
        st.rerun()

# =====================================================================
# 9. TELA: DASHBOARD
# =====================================================================
DISCIPLINAS_COMUNS = [
    "Matematica", "Portugues", "Historia", "Geografia", "Ciencias",
    "Biologia", "Fisica", "Quimica", "Ingles", "Espanhol", "Artes",
    "Educacao Fisica", "Filosofia", "Sociologia", "Ensino Religioso",
    "Informatica",
]

def render_calendario(planos):
    hoje = datetime.now()
    cal = st.session_state.get("dash_cal", hoje.strftime("%m/%Y"))
    mes, ano = parse_cal(cal)

    def mes_offset(delta):
        m, a = mes, ano
        m += delta
        if m > 12:
            m, a = 1, a + 1
        elif m < 1:
            m, a = 12, a - 1
        return m, a

    st.markdown('<div class="cal-widgets-marker"></div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1, 2.2, 1])
    voltar = c1.button("\u25c0", key="cal_prev", use_container_width=True)
    avancar = c3.button("\u25b6", key="cal_next", use_container_width=True)

    trocou_mes = False
    if voltar:
        mes, ano = mes_offset(-1)
        trocou_mes = True
    if avancar:
        mes, ano = mes_offset(1)
        trocou_mes = True
    if trocou_mes:
        st.session_state["dash_cal"] = f"{mes:02d}/{ano}"

    with c2:
        opcoes_pular = [f"{MESES_PT[m]} / {a}"
                        for a in (ano - 1, ano, ano + 1) for m in range(1, 13)]
        valor_atual = f"{MESES_PT[mes]} / {ano}"
        escolha = c2.selectbox("Mes / Ano", opcoes_pular,
                               index=opcoes_pular.index(valor_atual),
                               key=f"cal_ir_{mes}_{ano}", label_visibility="collapsed")
        if escolha != valor_atual:
            m_nome, a_str = escolha.rsplit(" / ", 1)
            mes, ano = MESES_PT.index(m_nome), int(a_str)
            trocou_mes = True
            st.session_state["dash_cal"] = f"{mes:02d}/{ano}"

    dias_cal = st.columns(7)
    for i, d in enumerate(DIAS_CURTO):
        dias_cal[i].markdown(
            f'<div style="text-align:center;font-size:.7rem;text-transform:uppercase;'
            f'letter-spacing:.08em;color:var(--cor-cinza);"> {d}</div>',
            unsafe_allow_html=True)

    dia_clicado = None
    matriz = calendar.monthcalendar(ano, mes)
    for semana in matriz:
        cols = st.columns(7)
        for i, dia in enumerate(semana):
            if dia == 0:
                continue
            data_str = f"{dia:02d}/{mes:02d}/{ano}"
            eh_hoje = (dia, mes, ano) == (hoje.day, hoje.month, hoje.year)
            tem_plano = len(planos.get(data_str, [])) > 0
            tipo = "primary" if tem_plano else "secondary"
            rotulo = f"**{dia}**" if eh_hoje else str(dia)
            if cols[i].button(rotulo, key=f"cal_d_{data_str}", type=tipo,
                              use_container_width=True):
                dia_clicado = data_str

    if trocou_mes and dia_clicado is None:
        return f"01/{mes:02d}/{ano}"
    return dia_clicado

@st.fragment
def fragmento_dashboard():
    hoje = datetime.now()
    planos = carregar_planos()
    grade = carregar_grade()
    anotacoes = carregar_anotacoes()

    dia_selecionada = st.session_state.get("dash_dia", hoje.strftime("%d/%m/%Y"))
    try:
        datetime.strptime(dia_selecionada, "%d/%m/%Y")
    except ValueError:
        dia_selecionada = hoje.strftime("%d/%m/%Y")
        st.session_state["dash_dia"] = dia_selecionada

    with st.container(key="dash_wrap"):
        with st.container(key="card_dash_cal"):
            novo_dia = render_calendario(planos)
            if novo_dia:
                st.session_state["dash_dia"] = novo_dia
                dia_selecionada = novo_dia

        with st.container(key="card_dash_grade"):
            st.markdown('<div class="card-titulo">Grade Semanal</div>', unsafe_allow_html=True)
            st.markdown(html_grade_mini(grade), unsafe_allow_html=True)

        with st.container(key="card_dash_mural"):
            topo = st.columns([3, 1])
            topo[0].markdown('<div class="card-titulo">Lembretes</div>', unsafe_allow_html=True)
            if topo[1].button("+ Novo", key="dash_novo_postit", use_container_width=True):
                dialog_postit()
            if not anotacoes:
                st.caption("Nenhum post-it. Crie um no botao + Novo.")
            for nota in anotacoes:
                cor = nota.get("cor") or "#fff3a3"
                txt = cor_texto_legivel(cor)
                st.markdown(
                    f'<div class="postit" style="background:{cor};color:{txt};">'
                    f'<div class="pt-titulo">\U0001f4cb {esc(nota.get("titulo", ""))}</div>'
                    f'<div class="pt-tag">{esc(nota.get("data", ""))}</div>'
                    f'<div class="pt-conteudo">{esc(nota.get("texto", ""))}</div>'
                    f'</div>',
                    unsafe_allow_html=True)
                b1, b2 = st.columns(2)
                if b1.button("Editar", key=f"dash_edit_{nota.get('id')}", use_container_width=True):
                    dialog_postit(nota)
                if b2.button("Excluir", key=f"dash_del_{nota.get('id')}", use_container_width=True):
                    anotacoes = [n for n in anotacoes if n.get("id") != nota.get("id")]
                    salvar_anotacoes(anotacoes)
                    st.rerun()

        with st.container(key="card_dash_agenda"):
            st.markdown(
                f'<div class="card-titulo">Agenda: {dia_selecionada}</div>',
                unsafe_allow_html=True)
            planos_dia = planos.get(dia_selecionada, [])
            if planos_dia:
                for i, plano in enumerate(planos_dia):
                    st.markdown(
                        f'<div class="dash-item">'
                        f'<div class="dash-item-titulo">'
                        f'{plano.get("horario","")} | {plano.get("turma","")}</div>'
                        f'<div class="dash-item-sub">'
                        f'{plano.get("disciplina","Geral")} - {plano.get("tema","Sem tema")}</div></div>',
                        unsafe_allow_html=True)
                    if st.button("Visualizar / Editar", key=f"dash_plano_{dia_selecionada}_{i}",
                                 use_container_width=True):
                        dialog_plano(dia_selecionada, i)
            else:
                st.caption("Nenhuma aula registrada neste dia.")

def tela_dashboard(config):
    hoje = datetime.now()

    dia_param = val_param("dia", "")
    try:
        datetime.strptime(dia_param, "%d/%m/%Y")
        st.session_state["dash_dia"] = dia_param
    except ValueError:
        if "dash_dia" not in st.session_state:
            st.session_state["dash_dia"] = hoje.strftime("%d/%m/%Y")

    cal_param = val_param("cal", "")
    if cal_param:
        st.session_state["dash_cal"] = cal_param
    if "dash_cal" not in st.session_state:
        st.session_state["dash_cal"] = hoje.strftime("%m/%Y")

    st.markdown(
        f"## Ola, {primeiro_nome_professor(config)}! :wave:")
    st.caption(f"Hoje e {hoje.strftime('%d/%m/%Y')}")

    fragmento_dashboard()

# =====================================================================
# 10. TELA: GRADE SEMANAL
# =====================================================================
def tela_grade_semanal():
    st.markdown("## Grade Semanal")
    grade = carregar_grade()
    max_aulas = carregar_config_grade()

    del_id = val_param("del")
    if del_id:
        try:
            del_id = int(del_id)
        except ValueError:
            del_id = None
        if del_id is not None and any(g.get("id") == del_id for g in grade):
            grade = [g for g in grade if g.get("id") != del_id]
            salvar_grade(grade)
            st.query_params.pop("del", None)
            st.rerun()

    c_cfg, c_add = st.columns([1, 2.4])
    with c_cfg:
        with st.container(key="card_grade_cfg"):
            nova_qtd = st.selectbox("Qtd. de aulas/dia:", [str(i) for i in range(1, 16)],
                                    index=max_aulas - 1 if 1 <= max_aulas <= 15 else 5)
            if st.button("Atualizar", use_container_width=True):
                salvar_config_grade(int(nova_qtd))
                st.rerun()

    with c_add:
        with st.container(key="card_grade_add"):
            st.markdown("**Adicionar aula a grade**")
            l1, l2 = st.columns([1, 1])
            nova_turma = l1.text_input("Turma", placeholder="Ex: 7o A", key="grade_turma")
            opcoes_disc = DISCIPLINAS_COMUNS + ["Outra disciplina..."]
            disc_sel = l2.selectbox("Disciplina", opcoes_disc,
                                    index=0, key="grade_disc_sel")
            nova_disc = ""
            if disc_sel == "Outra disciplina...":
                nova_disc = l2.text_input("Digite a disciplina:", key="grade_disc_outra")
            else:
                nova_disc = disc_sel
            dia = st.radio("Dia da semana", DIAS_UTEIS, horizontal=True)
            aulas_sel = st.multiselect(
                "Selecione as aulas", [f"{i}a Aula" for i in range(1, max_aulas + 1)],
                key="grade_aulas")
            if st.button("+ Adicionar", type="primary", use_container_width=True):
                if not nova_turma.strip() or not nova_disc.strip():
                    st.error("Preencha a turma e a disciplina.")
                elif not aulas_sel:
                    st.error("Marque pelo menos um horario.")
                else:
                    conflitos = [f"{a} (ocupada: {i['turma']})" for a in aulas_sel
                                 for i in grade if i["dia"] == dia and i["aula"] == a]
                    if conflitos:
                        st.warning("Conflito de horario:\n" + "\n".join(conflitos))
                    else:
                        novo_id = max([i.get("id", 0) for i in grade], default=0) + 1
                        for a in aulas_sel:
                            grade.append({"id": novo_id, "turma": nova_turma.strip(),
                                          "disciplina": nova_disc.strip(), "dia": dia, "aula": a})
                            novo_id += 1
                        salvar_grade(grade)
                        st.rerun()

    if not grade:
        st.info("Nenhum horario cadastrado. Comece adicionando turmas acima.")
        return

    def ordem_aula(aula):
        try:
            return int(str(aula).split("a")[0])
        except Exception:
            return 0

    with st.container(key="grade_dias"):
        colunas = st.columns(5)
        for i, dia in enumerate(DIAS_UTEIS):
            with colunas[i]:
                st.markdown(f"#### {dia.split('-')[0]}")
                itens = sorted([x for x in grade if x["dia"] == dia], key=lambda x: ordem_aula(x["aula"]))
                for item in itens:
                    st.markdown(
                        f'<div class="gcell">'
                        f'<a class="gcell-del" href="{link_para("Grade Semanal")}&del={item["id"]}" '
                        f'title="Remover este horario">&#10005;</a>'
                        f'<div class="gcell-aula">{item["aula"]}</div>'
                        f'<div class="gcell-turma">{item["turma"]}</div>'
                        f'<div class="gcell-disc">{item.get("disciplina","Geral")}</div></div>',
                        unsafe_allow_html=True)

# =====================================================================
# 11. TELA: TURMAS & ALUNOS
# =====================================================================
def tela_turmas():
    st.markdown("## Turmas & Alunos")
    grade = carregar_grade()
    dados_turmas = carregar_turmas()

    info_imp = st.session_state.pop("imp_info", None)
    if st.session_state.pop("imp_limpar", False):
        st.session_state.pop("imp_alunos", None)
    if info_imp:
        if info_imp[0] == "ok":
            st.success(info_imp[1])
        else:
            st.warning(info_imp[1])

    turmas_grade = sorted(list(set([i["turma"] for i in grade])))
    if not turmas_grade:
        st.error("Nenhuma turma cadastrada na Grade Semanal. Cadastre turmas la primeiro.")
        return

    if "turma_selecionada" not in st.session_state or st.session_state["turma_selecionada"] not in turmas_grade:
        st.session_state["turma_selecionada"] = turmas_grade[0]
    turma_atual = st.session_state["turma_selecionada"]

    col_menu, col_conteudo = st.columns([1, 2.6])

    with col_menu:
        with st.container(key="card_turmas_menu"):
            st.markdown("**Suas Turmas**")
            for turma in turmas_grade:
                qtd = len(dados_turmas.get(turma, []))
                ativa = turma == turma_atual
                if st.button(f"{turma}  ({qtd} alunos)", key=f"turma_{turma}",
                             type="primary" if ativa else "secondary",
                             use_container_width=True):
                    st.session_state["turma_selecionada"] = turma
                    st.rerun()

    with col_conteudo:
        st.markdown(f"### Alunos da turma: {turma_atual}")

        if st.session_state.get("confirmar_limpar_turma") == turma_atual:
            st.warning(f"Tem certeza que deseja remover TODOS os "
                       f"{len(dados_turmas.get(turma_atual, []))} alunos da turma {turma_atual}?")
            c1, c2 = st.columns(2)
            if c1.button("Sim, excluir todos", type="primary"):
                dados_turmas[turma_atual] = []
                salvar_turmas(dados_turmas)
                st.session_state["confirmar_limpar_turma"] = None
                st.rerun()
            if c2.button("Cancelar"):
                st.session_state["confirmar_limpar_turma"] = None
                st.rerun()

        arquivo = st.file_uploader("Importar alunos (Excel/CSV/TXT)", type=["xlsx", "xls", "csv", "txt"],
                                   key="imp_alunos")
        if arquivo is not None:
            resultado = processar_arquivo_alunos(arquivo, turma_atual, dados_turmas)
            st.session_state["imp_limpar"] = True
            st.session_state["imp_info"] = resultado
            st.rerun()

        l_add, l_limpar = st.columns([2.4, 1])
        novo_aluno = l_add.text_input("Nome do aluno", placeholder="Ex: Ana Clara", key="novo_aluno")
        c1, c2 = st.columns(2)
        if c1.button("+ Adicionar aluno", type="primary", use_container_width=True):
            nome = novo_aluno.strip()
            if not nome:
                st.error("Digite o nome do aluno.")
            elif nome in dados_turmas.get(turma_atual, []):
                st.warning("Este aluno ja esta cadastrado nesta turma.")
            else:
                dados_turmas.setdefault(turma_atual, []).append(nome)
                salvar_turmas(dados_turmas)
                st.rerun()
        if c2.button("Excluir todos", use_container_width=True):
            st.session_state["confirmar_limpar_turma"] = turma_atual
            st.rerun()

        alunos = sorted(dados_turmas.get(turma_atual, []))
        if not alunos:
            st.caption("Nenhum aluno cadastrado nesta turma ainda.")
        for i, aluno in enumerate(alunos, 1):
            c1, c2 = st.columns([6, 1])
            c1.markdown(f"**{i}. {aluno}**")
            if c2.button("x", key=f"aluno_{turma_atual}_{aluno}"):
                dados_turmas[turma_atual] = [a for a in dados_turmas[turma_atual] if a != aluno]
                salvar_turmas(dados_turmas)
                st.rerun()

def processar_arquivo_alunos(arquivo, turma_atual, dados_turmas):
    nome_arq = arquivo.name.lower()
    novos = []
    try:
        if nome_arq.endswith((".xlsx", ".xls")):
            wb = openpyxl.load_workbook(io.BytesIO(arquivo.read()), data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if row and row[0]:
                    nome = str(row[0]).strip()
                    if nome and nome.lower() not in ["nome", "alunos", "estudante", "aluno"]:
                        novos.append(nome)
        else:
            conteudo = arquivo.read()
            texto = None
            for enc in ("utf-8", "latin-1"):
                try:
                    texto = conteudo.decode(enc)
                    break
                except Exception:
                    continue
            if texto is None:
                texto = conteudo.decode("utf-8", errors="ignore")
            leitor = csv.reader(io.StringIO(texto), delimiter=";")
            for linha in leitor:
                if linha:
                    nome = linha[0].strip()
                    if nome and nome.lower() not in ["nome", "alunos", "estudante", "aluno"]:
                        novos.append(nome)
            if len(novos) <= 1:
                novos = [l.strip() for l in texto.splitlines()
                         if l.strip() and l.strip().lower() not in ["nome", "alunos", "estudante", "aluno"]]
    except Exception as e:
        return ("erro", f"Nao foi possivel ler o arquivo: {e}")

    if not novos:
        return ("vazio", "O arquivo parece estar vazio ou sem nomes na primeira coluna.")

    dados_turmas.setdefault(turma_atual, [])
    adicionados = 0
    for nome in novos:
        if nome not in dados_turmas[turma_atual]:
            dados_turmas[turma_atual].append(nome)
            adicionados += 1
    salvar_turmas(dados_turmas)
    return ("ok", f"{adicionados} aluno(s) importado(s) para a turma {turma_atual}!")

# =====================================================================
# 12. TELA: CENTRAL DE PLANOS
# =====================================================================
def html_calendario_planos(planos, grade, ano, mes, turma_filtro, mostrar_audit, dark):
    dias_completos = DIAS_COMPLETOS
    base_cor = "#333333" if dark else "#d6d6d6"
    base_txt = "#ffffff" if dark else "#000000"

    linhas = ['<div class="cal-scroll"><table class="cal-table"><tr>']
    for d in DIAS_CURTO:
        linhas.append(f"<th>{d}</th>")
    linhas.append("</tr>")

    matriz = calendar.monthcalendar(ano, mes)
    for semana in matriz:
        linhas.append("<tr>")
        for coluna, dia in enumerate(semana):
            if dia == 0:
                linhas.append('<td style="border:none;padding:4px;"></td>')
                continue
            data_str = f"{dia:02d}/{mes:02d}/{ano}"
            nome_dia = dias_completos[coluna]
            aulas_esperadas = [g for g in grade if g["dia"] == nome_dia]
            if turma_filtro != "Todas as Turmas":
                aulas_esperadas = [g for g in aulas_esperadas if g["turma"] == turma_filtro]
            qtd_esperada = len(aulas_esperadas)

            planos_dia = planos.get(data_str, [])
            if turma_filtro != "Todas as Turmas":
                planos_dia = [p for p in planos_dia if p["turma"] == turma_filtro]
            qtd_registrada = len(planos_dia)

            if mostrar_audit:
                if qtd_esperada == 0:
                    cor, txt = base_cor, "#a0a0a0"
                elif qtd_registrada == 0:
                    cor, txt = "#dc3545", "#ffffff"
                elif qtd_registrada < qtd_esperada:
                    cor, txt = "#ffc107", "#000000"
                else:
                    cor, txt = "#28a745", "#ffffff"
            else:
                if qtd_registrada > 0:
                    cor, txt = "#007bff", "#ffffff"
                else:
                    cor, txt = base_cor, base_txt

            linhas.append(
                f'<td style="background:{cor};color:{txt};border:1px solid rgba(128,128,128,0.25);'
                f'text-align:center;font-weight:600;padding:8px 0;border-radius:8px;">{dia}</td>')
        linhas.append("</tr>")
    linhas.append("</table></div>")
    return "".join(linhas)

CAMPOS_PLANO = [
    ("Metodologia / Atividades", "metodologia"),
    ("Objetivos Especificos", "objetivos"),
    ("Procedimentos", "procedimentos"),
    ("Habilidade(s) da BNCC", "habilidades"),
    ("Competencia Geral", "comp_geral"),
    ("Competencias Especificas", "comp_especifica"),
    ("Recursos Necessarios", "recursos"),
    ("Avaliacao", "avaliacao"),
    ("Observacoes", "observacoes"),
]

def tela_central_planos(config):
    st.markdown("## Central de Planos")
    grade = carregar_grade()
    if not grade:
        st.error("Sua grade esta vazia. Cadastre turmas na Grade Semanal primeiro.")
        return

    tab_criar, tab_listar = st.tabs(["Distribuir Novo Plano", "Planos deste Mes"])

    with tab_criar:
        with st.container(key="card_plano_criar"):
            tema = st.text_input("Tema / Conteudo da aula", placeholder="Ex: Funcoes do 1o grau")
            c1, c2 = st.columns(2)
            data_inicio = c1.text_input("Data de inicio (DD/MM/AAAA)",
                                        value=datetime.now().strftime("%d/%m/%Y"))
            carga = c2.selectbox("Carga horaria (aulas)", [f"{i} Aulas" for i in range(1, 11)])
            st.markdown("**Selecione as turmas e disciplinas para distribuir:**")
            turmas_disciplinas = []
            for item in grade:
                td = f"{item['turma']} - {item.get('disciplina', 'Geral')}"
                if td not in turmas_disciplinas:
                    turmas_disciplinas.append(td)
            td_selecionadas = st.multiselect("Turmas / Disciplinas", turmas_disciplinas)

            aba_m, aba_b, aba_r = st.tabs(["Metodologia & Objetivos", "Estrutura BNCC", "Recursos & Avaliacao"])
            mapa = {"metodologia": 0, "objetivos": 0, "procedimentos": 0,
                    "habilidades": 1, "comp_geral": 1, "comp_especifica": 1,
                    "recursos": 2, "avaliacao": 2, "observacoes": 2}
            valores_extras = {}
            for rotulo, chave in CAMPOS_PLANO:
                with (aba_m if mapa[chave] == 0 else aba_b if mapa[chave] == 1 else aba_r):
                    valores_extras[chave] = st.text_area(rotulo, key=f"plano_{chave}")

            if st.button("Distribuir Plano Automaticamente", type="primary",
                         use_container_width=True):
                if not tema.strip():
                    st.error("O campo Tema e obrigatorio.")
                elif not valores_extras["metodologia"].strip():
                    st.error("O campo Metodologia e obrigatorio.")
                elif not td_selecionadas:
                    st.error("Marque pelo menos uma turma/disciplina.")
                else:
                    try:
                        data_obj = datetime.strptime(data_inicio.strip(), "%d/%m/%Y")
                    except ValueError:
                        st.error("Formato de data invalido. Use DD/MM/AAAA.")
                    else:
                        carga_n = int(carga.split()[0])
                        resumo = distribuir_planos(planos=carregar_planos(), grade=grade,
                                                   td_selecionadas=td_selecionadas,
                                                   tema=tema.strip(), data_inicio=data_obj,
                                                   carga_horaria=carga_n,
                                                   dados_extras={k: v.strip() for k, v in valores_extras.items()})
                        if resumo:
                            st.success("Plano distribuido com sucesso!\n\n" + "\n".join(resumo))
                        else:
                            st.warning("Nao foi possivel encontrar horarios na grade para essas turmas.")

    with tab_listar:
        montar_aba_listar_planos(grade, config)

def distribuir_planos(planos, grade, td_selecionadas, tema, data_inicio,
                      carga_horaria, dados_extras):
    mapa_dias = {d: i for i, d in enumerate(DIAS_COMPLETOS)}
    resumo = []
    plano_modificado = False

    for td in td_selecionadas:
        partes = td.split(" - ")
        turma_nome = partes[0]
        disc_nome = partes[1] if len(partes) > 1 else "Geral"

        aulas_restantes = carga_horaria
        aulas_turma = [i for i in grade
                       if i["turma"] == turma_nome and i.get("disciplina", "Geral") == disc_nome]
        if not aulas_turma:
            continue

        dia_atual = data_inicio
        limite_busca = 30
        while aulas_restantes > 0 and limite_busca > 0:
            dia_semana_teste = dia_atual.weekday()
            aulas_neste_dia = [a for a in aulas_turma if mapa_dias.get(a["dia"]) == dia_semana_teste]
            aulas_neste_dia.sort(key=lambda x: x.get("aula", ""))

            for aula_obj in aulas_neste_dia:
                if aulas_restantes <= 0:
                    break
                str_data_final = dia_atual.strftime("%d/%m/%Y")
                if str_data_final not in planos:
                    planos[str_data_final] = []
                novo_plano = {
                    "turma": turma_nome, "disciplina": disc_nome,
                    "horario": aula_obj["aula"], "tema": tema
                }
                novo_plano.update(dados_extras)
                planos[str_data_final].append(novo_plano)
                resumo.append(f"{td}: {str_data_final} ({aula_obj['aula']})")
                aulas_restantes -= 1
                plano_modificado = True

            dia_atual += timedelta(days=1)
            limite_busca -= 1

    if plano_modificado:
        salvar_planos(planos)
    return resumo

def montar_aba_listar_planos(grade, config):
    hoje = datetime.now()
    pm = val_param("pm", hoje.strftime("%m/%Y"))
    try:
        mes, ano = [int(x) for x in pm.split("/")]
    except Exception:
        mes, ano = hoje.month, hoje.year
    if mes < 1 or mes > 12:
        mes, ano = hoje.month, hoje.year

    def href_pm(novo_mes, novo_ano):
        return f"?pagina=Central de Planos&pm={novo_mes}/{novo_ano}"

    def mes_offset(delta):
        m, a = mes, ano
        m += delta
        if m > 12:
            m, a = 1, a + 1
        elif m < 1:
            m, a = 12, a - 1
        return m, a

    c1, c2, c3 = st.columns([1, 2.2, 1])
    m_prev, m_next = mes_offset(-1), mes_offset(1)
    c1.markdown(f'<a href="{href_pm(m_prev[0], m_prev[1])}" style="text-decoration:none;'
                f'background:#6c757d;color:#fff;padding:8px 16px;border-radius:8px;'
                f'font-weight:600;">&#9664; Anterior</a>', unsafe_allow_html=True)
    c2.markdown(f"### {MESES_PT[mes]} de {ano}", unsafe_allow_html=True)
    c3.markdown(f'<div style="text-align:right;"><a href="{href_pm(m_next[0], m_next[1])}" '
                f'style="text-decoration:none;background:#6c757d;color:#fff;padding:8px 16px;'
                f'border-radius:8px;font-weight:600;">Proximo &#9654;</a></div>', unsafe_allow_html=True)

    planos = carregar_planos()
    turmas_unicas = ["Todas as Turmas"] + sorted(list(set([i["turma"] for i in grade])))
    turma_filtro = st.selectbox("Filtrar por turma", turmas_unicas)
    mostrar_audit = st.toggle("Mostrar Auditoria (Faltas de Planejamento)", value=False)

    dark = config.get("aparencia", "System") == "Dark"
    legenda = ("Modo Auditoria: verde = 100% planejado | amarelo = parcial | "
               "vermelho = falta plano | cinza = sem aula" if mostrar_audit else
               "Modo Normal: azul = aulas planejadas neste dia | cinza = sem registros")
    st.caption(legenda)

    with st.container(key="card_planos_cal"):
        st.markdown(html_calendario_planos(planos, grade, ano, mes, turma_filtro, mostrar_audit, dark),
                    unsafe_allow_html=True)

    if not mostrar_audit:
        datas_mes = []
        for data_str in planos.keys():
            try:
                d = datetime.strptime(data_str, "%d/%m/%Y")
                if d.month == mes and d.year == ano:
                    datas_mes.append((d, data_str))
            except Exception:
                continue
        datas_mes.sort()
        encontrou = False
        for data_obj, data_str in datas_mes:
            for i, plano in enumerate(planos[data_str]):
                if turma_filtro != "Todas as Turmas" and plano.get("turma") != turma_filtro:
                    continue
                encontrou = True
                st.markdown(
                    f'<div class="dash-item">'
                    f'<div class="dash-item-titulo">{data_str} - {DIAS_COMPLETOS[data_obj.weekday()]} | '
                    f'{plano.get("horario","")} | {plano.get("turma","")} ({plano.get("disciplina","Geral")})</div>'
                    f'<div class="dash-item-sub">Tema: {plano.get("tema","")}</div></div>',
                    unsafe_allow_html=True)
                if st.button("Visualizar / Editar", key=f"plist_{data_str}_{i}"):
                    dialog_plano(data_str, i)
        if not encontrou:
            st.caption("Nenhum plano registrado para os filtros neste mes.")
    else:
        pendencias = 0
        ultimo_dia = calendar.monthrange(ano, mes)[1]
        for dia in range(1, ultimo_dia + 1):
            data_str = f"{dia:02d}/{mes:02d}/{ano}"
            data_obj = datetime(ano, mes, dia)
            nome_dia = DIAS_COMPLETOS[data_obj.weekday()]
            aulas_esperadas = [g for g in grade if g["dia"] == nome_dia]
            if turma_filtro != "Todas as Turmas":
                aulas_esperadas = [g for g in aulas_esperadas if g["turma"] == turma_filtro]
            if not aulas_esperadas:
                continue
            planos_dia = planos.get(data_str, [])
            faltas = []
            for esperada in aulas_esperadas:
                registrado = any(p.get("turma") == esperada["turma"] and p.get("horario") == esperada["aula"]
                                 for p in planos_dia)
                if not registrado:
                    faltas.append(esperada)
            if faltas:
                pendencias += 1
                st.markdown(f"**PENDENCIA: {data_str} - {nome_dia}**")
                for falta in faltas:
                    st.markdown(
                        f'<div class="pend-item">'
                        f'Falta plano para a {falta["aula"]} | Turma: {falta["turma"]} '
                        f'({falta.get("disciplina","Geral")})</div>', unsafe_allow_html=True)
        if pendencias == 0:
            st.success("Parabens! Tudo 100% planejado para os filtros selecionados.")

# =====================================================================
# 13. TELA: ANOTACOES
# =====================================================================
def tela_anotacoes():
    st.markdown("## Lembretes")
    anotacoes = carregar_anotacoes()
    grade = carregar_grade()

    c1, c2 = st.columns([3, 1])
    turmas = ["Todas as Notas"] + sorted(list(set([i["turma"] for i in grade])))
    filtro = c1.selectbox("Filtrar mural", turmas)
    if c2.button("+ Novo Post-it", type="primary", use_container_width=True):
        dialog_postit()

    notas = anotacoes
    if filtro != "Todas as Notas":
        notas = [n for n in anotacoes if n.get("turma") == filtro or n.get("turma") == "Geral"]

    if not notas:
        st.caption("Nenhuma anotacao por aqui. Clique em Novo Post-it acima!")
        return

    colunas = st.columns(3)
    for i, nota in enumerate(notas):
        with colunas[i % 3]:
            cor = nota.get("cor", "#fff3a3")
            txt = cor_texto_legivel(cor)
            st.markdown(
                f'<div class="postit" style="background:{cor};color:{txt};min-height:120px;">'
                f'<div class="pt-titulo">{nota.get("titulo","Sem titulo")}</div>'
                f'<div class="pt-tag">#{nota.get("turma","Geral")} | {nota.get("data_criacao","")}</div>'
                f'<div class="pt-conteudo">{nota.get("conteudo","")}</div></div>',
                unsafe_allow_html=True)
            b1, b2 = st.columns(2)
            if b1.button("Editar", key=f"anot_edit_{nota.get('id')}", use_container_width=True):
                dialog_postit(nota)
            if b2.button("Excluir", key=f"anot_del_{nota.get('id')}", use_container_width=True):
                anotacoes = [n for n in anotacoes if n.get("id") != nota.get("id")]
                salvar_anotacoes(anotacoes)
                st.rerun()

# =====================================================================
# 14. TELA: NOTAS E ESTATISTICAS
# =====================================================================
def nivel_info(perc):
    if perc < 50:
        return "Muito Critico", "#dc3545", "#ffffff"
    elif perc < 70:
        return "Critico", "#ffc107", "#000000"
    elif perc < 90:
        return "Intermediario", "#66bb6a", "#000000"
    else:
        return "Adequado", "#1b5e20", "#ffffff"

def resolver_descritor(texto):
    m = re.search(r"\d+", texto.strip())
    if m:
        chave = f"D{int(m.group())}"
        if chave in DESCRITORES_MAT:
            return chave
    return "Sem Descritor"

def tela_notas():
    st.markdown("## Central de Notas e Estatisticas")
    avaliacoes = carregar_avaliacoes()
    dados_turmas = carregar_turmas()

    tab_criar, tab_corrigir, tab_stats = st.tabs(
        ["1. Criar Avaliacao & Gabarito", "2. Larcar Notas / Corrigir", "3. Estatisticas por Descritor"])

    # ---------------- TAB 1: CRIAR ----------------
    with tab_criar:
        turmas_cad = sorted(list(dados_turmas.keys()))
        if not turmas_cad:
            st.error("Voce nao tem turmas com alunos. Cadastre em Turmas & Alunos primeiro.")
        else:
            c1, c2, c3 = st.columns(3)
            turmas_sel = c1.multiselect("Turmas que farao a prova", turmas_cad)
            titulo = c2.text_input("Titulo da prova", placeholder="Ex: Prova Bimestral de Matematica")
            qtd = c3.number_input("Qtd. de questoes", min_value=1, max_value=50, value=10, step=1)

            if st.button("Gerar Prova", type="primary"):
                st.session_state["gab_qtd"] = int(qtd)
                st.session_state["gab_rows"] = [{"resposta": "A", "descritor": ""}
                                                for _ in range(int(qtd))]

            st.markdown("**Gabarito e descritores (digite o numero, ex: 2, 02, D02):**")
            linhas = st.session_state.get("gab_rows", [])
            for i, row in enumerate(linhas):
                r1, r2, r3 = st.columns([1, 1, 3])
                chave_r = f"gab_{i}"
                resp = r1.selectbox(f"Q{i + 1:02d} - Resposta", ["A", "B", "C", "D", "E"],
                                    index=["A", "B", "C", "D", "E"].index(row.get("resposta", "A")),
                                    key=f"{chave_r}_resp")
                desc = r2.text_input(f"Descritor Q{i + 1:02d}", value=row.get("descritor", ""),
                                     placeholder="Ex: D02", key=f"{chave_r}_desc")
                hint = ""
                if desc.strip():
                    chave_d = resolver_descritor(desc)
                    if chave_d != "Sem Descritor":
                        hint = f":blue[{DESCRITORES_MAT[chave_d]}]"
                    else:
                        hint = ":red[Descritor nao encontrado]"
                r3.caption(hint if hint else " ")
                row["resposta"] = resp
                row["descritor"] = desc

            if st.button("Salvar Gabarito para Turmas Selecionadas", type="primary",
                         use_container_width=True):
                if not turmas_sel:
                    st.error("Selecione pelo menos uma turma.")
                elif not linhas:
                    st.error("Clique em 'Gerar Prova' para gerar os campos do gabarito.")
                else:
                    gabarito_final = [
                        {"questao": i + 1,
                         "resposta_correta": row["resposta"],
                         "descritor": resolver_descritor(row["descritor"])}
                        for i, row in enumerate(linhas)]
                    criadas = 0
                    for turma in turmas_sel:
                        avaliacoes.append({
                            "id": max([a.get("id", 0) for a in avaliacoes], default=0) + 1,
                            "titulo": titulo.strip() or "Avaliacao sem Titulo",
                            "turma": turma,
                            "data": datetime.now().strftime("%d/%m/%Y"),
                            "gabarito": gabarito_final,
                            "notas_alunos": {}
                        })
                        criadas += 1
                    salvar_avaliacoes(avaliacoes)
                    st.success(f"{criadas} avaliacao(oes) salva(s) com sucesso!")
                    st.session_state.pop("gab_rows", None)

    # ---------------- TAB 2: CORRIGIR ----------------
    with tab_corrigir:
        if not avaliacoes:
            st.info("Nenhuma avaliacao cadastrada.")
        else:
            titulos_av = [f"[{a['turma']}] {a['titulo']} ({a['data']})" for a in avaliacoes]
            c1, c2 = st.columns([2.5, 1.5])
            av_escolhida = c1.selectbox("Selecione a avaliacao", titulos_av, key="av_corrigir")
            modo = c2.radio("Modo de correcao", ["Rapido (Texto)", "Detalhado (Blocos)"],
                            horizontal=True)

            idx = titulos_av.index(av_escolhida)
            av = avaliacoes[idx]
            turma = av["turma"]
            alunos = sorted(dados_turmas.get(turma, []))
            gabarito = av["gabarito"]
            qtd_q = len(gabarito)

            if not alunos:
                st.warning("Esta turma nao tem alunos cadastrados.")
            else:
                if modo == "Rapido (Texto)":
                    st.caption("Dica: digite as respostas seguidas (Ex: ADCBE) na caixa.")
                else:
                    st.caption("Modo detalhado: preencha a letra marcada em cada bloco.")

                with st.form("form_correcao"):
                    dados_entrada = []
                    for aluno in alunos:
                        st.markdown(f"**{aluno}**")
                        dados_exist = av["notas_alunos"].get(aluno, {})
                        if modo == "Rapido (Texto)":
                            c1, c2 = st.columns([3, 1])
                            resp_txt = c1.text_input(
                                f"Respostas ({qtd_q} questoes)",
                                value=dados_exist.get("respostas_dadas", ""),
                                key=f"corr_resp_{aluno}", label_visibility="collapsed")
                            nota = c2.text_input("Nota", value=str(dados_exist.get("nota_final", "")),
                                                 key=f"corr_nota_{aluno}",
                                                 placeholder="auto")
                            dados_entrada.append((aluno, "rapido", resp_txt, nota, None))
                        else:
                            notas_exist = dados_exist.get("respostas_dadas", "")
                            blocos = []
                            cols = st.columns(min(qtd_q, 5))
                            for qi in range(qtd_q):
                                val = notas_exist[qi] if qi < len(notas_exist) else ""
                                with cols[qi % min(qtd_q, 5)]:
                                    inp = st.text_input(f"Q{qi + 1}", value=val,
                                                        key=f"corr_{aluno}_q{qi}",
                                                        max_chars=1)
                                    blocos.append(inp)
                            nota = st.text_input("Nota final", key=f"corr_nota_d_{aluno}",
                                                 value=str(dados_exist.get("nota_final", "")),
                                                 placeholder="auto")
                            dados_entrada.append((aluno, "detalhado", None, nota, blocos))

                    enviado = st.form_submit_button("Processar Correcao Automatica e Salvar",
                                                    type="primary", use_container_width=True)
                    if enviado:
                        processar_correcao(av, dados_entrada, gabarito, avaliacoes, idx)

    # ---------------- TAB 3: ESTATISTICAS ----------------
    with tab_stats:
        montar_aba_stats(avaliacoes)

# ---------------- correcao ----------------
def processar_correcao(av, dados_entrada, gabarito, avaliacoes, idx):
    qtd_q = len(gabarito)
    for aluno, tipo, resp_txt, nota_txt, blocos in dados_entrada:
        if tipo == "rapido":
            resp_dadas = resp_txt.strip().upper()
        else:
            resp_dadas = "".join([(b.strip().upper() or " ") for b in blocos])

        nota_manual = nota_txt.strip()
        acertos_descritores = {}
        nota_calculada = 0

        if resp_dadas.strip():
            acertos = 0
            for i, q in enumerate(gabarito):
                descritor = q["descritor"]
                if descritor not in acertos_descritores:
                    acertos_descritores[descritor] = {"corretas": 0, "total": 0}
                acertos_descritores[descritor]["total"] += 1
                letra_aluno = resp_dadas[i] if i < len(resp_dadas) else " "
                if letra_aluno == q["resposta_correta"].upper():
                    acertos += 1
                    acertos_descritores[descritor]["corretas"] += 1
            nota_calculada = round((acertos / qtd_q) * 10, 1)

        if resp_dadas.strip() or nota_manual:
            av["notas_alunos"][aluno] = {
                "respostas_dadas": resp_dadas,
                "nota_final": float(nota_manual) if nota_manual else nota_calculada,
                "descritores": acertos_descritores
            }
    avaliacoes[idx] = av
    salvar_avaliacoes(avaliacoes)
    st.success("Correcao automatica finalizada e notas salvas!")
    st.rerun()

# =====================================================================
# 15. ESTATISTICAS POR DESCRITOR
# =====================================================================
def montar_aba_stats(avaliacoes):
    if not avaliacoes:
        st.info("Nenhuma avaliacao cadastrada.")
        return

    titulos_av = [f"[{a['turma']}] {a['titulo']} ({a['data']})" for a in avaliacoes]
    av_escolhida = st.selectbox("Avaliacao para estatisticas", titulos_av, key="av_stats")

    idx = titulos_av.index(av_escolhida)
    av = avaliacoes[idx]
    notas = av["notas_alunos"]
    gabarito = av["gabarito"]
    qtd_questoes = len(gabarito)

    gerar = st.button("Gerar Relatorio", type="primary")
    if not gerar:
        return

    if not notas:
        st.warning("Nenhuma nota lancada para gerar estatisticas.")
        return

    media_turma = sum([n["nota_final"] for n in notas.values()]) / len(notas)
    distribuicao = {"Muito Critico": 0, "Critico": 0, "Intermediario": 0, "Adequado": 0}
    for dados in notas.values():
        perc_aluno = (dados["nota_final"] / 10) * 100
        nivel, _, _ = nivel_info(perc_aluno)
        distribuicao[nivel] += 1

    st.markdown(f"### Resumo da turma: {av['turma']}")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Media da Turma", f"{media_turma:.1f}")
    c2.metric("Adequado (90-100%)", distribuicao["Adequado"])
    c3.metric("Intermediario (70-89%)", distribuicao["Intermediario"])
    c4.metric("Critico (50-69%)", distribuicao["Critico"])
    c5.metric("Muito Critico (<50%)", distribuicao["Muito Critico"])

    st.markdown("### Desempenho por Descritor / Habilidade")
    consolidado = {}
    for aluno, dados in notas.items():
        for desc, val in dados.get("descritores", {}).items():
            if desc not in consolidado:
                consolidado[desc] = {"corretas": 0, "total": 0}
            consolidado[desc]["corretas"] += val["corretas"]
            consolidado[desc]["total"] += val["total"]

    if consolidado:
        for desc, val in sorted(consolidado.items()):
            if val["total"] == 0:
                continue
            perc = (val["corretas"] / val["total"]) * 100
            nome_nivel, cor_barra, _ = nivel_info(perc)
            nome_desc = DESCRITORES_MAT.get(desc, "")
            if nome_desc:
                st.caption(f"{desc} - {nome_desc}")
            st.progress(perc / 100, text=f"{desc}: {val['corretas']}/{val['total']} acertos "
                                         f"({perc:.1f}% - {nome_nivel})")
    else:
        st.caption("Sem dados de descritores.")

    st.markdown("### Detalhamento por Aluno")
    linhas_tab = []
    for aluno, dados in sorted(notas.items()):
        perc = (dados["nota_final"] / 10) * 100
        nivel, _, _ = nivel_info(perc)
        acertos = round((dados["nota_final"] / 10) * qtd_questoes)
        linhas_tab.append({"Aluno": aluno, "Acertos": f"{acertos}/{qtd_questoes}",
                           "Nota (0-10)": f"{dados['nota_final']:.1f}",
                           "Nivel": f"{perc:.0f}% - {nivel}"})
    st.dataframe(linhas_tab, use_container_width=True, hide_index=True)

    excel = gerar_excel_avaliacao(av)
    st.download_button(
        "Exportar Relatorio para Excel",
        data=excel,
        file_name=f"Relatorio_{av['turma']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")

# =====================================================================
# 16. TELA: CADASTRAR NOVA QUESTAO
# =====================================================================
def tela_cadastrar():
    st.markdown("## Cadastrar Nova Questao")
    banco = carregar_banco()

    info_cad = st.session_state.pop("cad_info", None)
    if st.session_state.pop("cad_limpar_img", False):
        st.session_state.pop("img_nova_questao", None)
    if info_cad:
        if info_cad[0] == "ok":
            st.success(info_cad[1])
        else:
            st.error(info_cad[1])

    with st.form("form_nova_questao"):
        enunciado = st.text_area("1. Enunciado / Texto de Contexto Inicial:", height=110,
                                 placeholder="Cole aqui o texto de contexto da questao...")
        imagem = st.file_uploader("2. Imagem de Apoio (Opcional):",
                                  type=["png", "jpg", "jpeg", "bmp", "webp"], key="img_nova_questao")
        pergunta = st.text_area("3. Pergunta Direta / Comando:", height=70,
                                placeholder="Ex: Qual e o valor de x?")

        st.markdown("**4. Alternativas (deixe em branco se for discursiva):**")
        c_alt = st.columns(2)
        alt_a = c_alt[0].text_input("A)", placeholder="Texto da alternativa A")
        alt_b = c_alt[1].text_input("B)", placeholder="Texto da alternativa B")
        alt_c = c_alt[0].text_input("C)", placeholder="Texto da alternativa C")
        alt_d = c_alt[1].text_input("D)", placeholder="Texto da alternativa D")

        c2 = st.columns(3)
        gabarito = c2[0].selectbox("5. Gabarito Correto:", ["A", "B", "C", "D", "Discursiva"])
        dificuldade = c2[1].selectbox("6. Nivel de Dificuldade:", ["Facil", "Medio", "Dificil"])
        tema = c2[2].text_input("7. Tema / Descritor SAEB:",
                                placeholder="Ex: D15 - Resolver problema...")
        salvar = st.form_submit_button("Salvar Questao", type="primary",
                                       use_container_width=True)

    if salvar:
        if not enunciado.strip() or not pergunta.strip() or not tema.strip():
            st.error("Por favor, preencha o Enunciado, a Pergunta e o Tema!")
        else:
            alternativas = {}
            for letra, texto in [("A", alt_a), ("B", alt_b), ("C", alt_c), ("D", alt_d)]:
                if texto.strip():
                    alternativas[letra] = texto.strip()

            novo_id = max([q.get("id", 0) for q in banco], default=0) + 1
            caminho_final_img = ""
            if imagem is not None:
                nome_arq = imagem.name.replace(" ", "_")
                caminho_final_img = salvar_imagem_usuario(nome_arq, imagem.getbuffer())

            banco.append({
                "id": novo_id,
                "tema": tema.strip(),
                "dificuldade": dificuldade,
                "enunciado": enunciado.strip(),
                "imagem": caminho_final_img,
                "pergunta_direta": pergunta.strip(),
                "alternativas": alternativas,
                "gabarito": gabarito
            })
            salvar_banco(banco)
            st.session_state["cad_limpar_img"] = True
            st.session_state["cad_info"] = ("ok",
                                            f"Questao {novo_id} salva com sucesso no banco de dados!")
            st.rerun()

# =====================================================================
# 17. TELA: IMPORTAR LOTE (IA)
# =====================================================================
def tela_importar():
    st.markdown("## Importar Lote de Questoes (IA)")
    st.caption("Cole na caixa abaixo o codigo JSON gerado pela Inteligencia Artificial. "
               "Ele deve conter uma lista de questoes, por exemplo:\n\n"
               '`[ { "tema": "...", "enunciado": "...", "pergunta_direta": "..." } ]`')

    texto_json = st.text_area("Codigo JSON do lote:", height=320,
                              key="importar_json", placeholder='[ { "tema": "...", ... } ]')

    if st.button("Processar e Salvar Lote", type="primary", use_container_width=True):
        if not texto_json.strip():
            st.warning("Por favor, cole o codigo gerado pela IA antes de processar.")
            return
        try:
            novas_questoes = json.loads(texto_json)
            if not isinstance(novas_questoes, list):
                st.error("O formato esta incorreto. Certifique-se de que e uma lista [ {...}, {...} ].")
                return

            banco = carregar_banco()
            sucesso = 0
            maior_id = max([q.get("id", 0) for q in banco], default=0)
            for q in novas_questoes:
                if not isinstance(q, dict):
                    continue
                if "enunciado" in q and "pergunta_direta" in q and "tema" in q:
                    maior_id += 1
                    banco.append({
                        "id": maior_id,
                        "tema": str(q.get("tema", "")).strip(),
                        "dificuldade": q.get("dificuldade", "Medio"),
                        "enunciado": q.get("enunciado", ""),
                        "imagem": q.get("imagem", ""),
                        "pergunta_direta": q.get("pergunta_direta", ""),
                        "alternativas": q.get("alternativas", {}),
                        "gabarito": q.get("gabarito", "")
                    })
                    sucesso += 1

            if sucesso > 0:
                salvar_banco(banco)
                st.success(f"{sucesso} questao(oes) foram importadas e adicionadas ao seu catalogo!")
            else:
                st.warning("Nenhuma questao valida encontrada no codigo inserido. Verifique a estrutura.")
        except json.JSONDecodeError as e:
            st.error(f"Erro de formatacao no codigo colado. Verifique se copiou o texto completo.\n\nDetalhe tecnico: {e}")

# =====================================================================
# 18. TELA: CATALOGO DE QUESTOES
# =====================================================================
@st.dialog("Editar Questao")
def dialog_editar_questao(id_q):
    banco = carregar_banco()
    q = next((item for item in banco if item.get("id") == id_q), None)
    if not q:
        st.info("Questao nao encontrada.")
        return
    novo_enun = st.text_area("Enunciado:", value=q.get("enunciado", ""), height=110)
    nova_perg = st.text_area("Pergunta Direta:", value=q.get("pergunta_direta", ""), height=60)
    novo_tema = st.text_input("Tema / Descritor:", value=q.get("tema", ""))
    novo_gab = st.text_input("Gabarito:", value=q.get("gabarito", ""))
    if st.button("Salvar Alteracoes", type="primary", use_container_width=True):
        q["enunciado"] = novo_enun.strip()
        q["pergunta_direta"] = nova_perg.strip()
        q["tema"] = novo_tema.strip()
        q["gabarito"] = novo_gab.strip()
        salvar_banco(banco)
        st.rerun()

@st.dialog("Confirmar Exclusao")
def dialog_excluir_questao(id_q):
    st.warning("Tem certeza que deseja excluir esta questao permanentemente?")
    c1, c2 = st.columns(2)
    if c1.button("Sim, excluir", type="primary", use_container_width=True):
        banco = carregar_banco()
        salvar_banco([q for q in banco if q.get("id") != id_q])
        st.rerun()
    if c2.button("Cancelar", use_container_width=True):
        st.rerun()

def tela_catalogo():
    st.markdown("## Catalogo de Questoes")
    banco = carregar_banco()
    if not banco:
        st.info("O banco de questoes esta vazio! Cadastre ou importe questoes.")
        return

    c1, c2 = st.columns([3, 1])
    termo = c1.text_input("Buscar por palavra-chave...", key="cat_busca")
    dificuldade_f = c2.selectbox("Nivel:", ["Todos", "Facil", "Medio", "Dificil"], key="cat_dif")

    t = termo.strip().lower()
    questoes = []
    for q in banco:
        if dificuldade_f != "Todos" and q.get("dificuldade", "") != dificuldade_f:
            continue
        if t and t not in q.get("enunciado", "").lower() \
                and t not in q.get("pergunta_direta", "").lower() \
                and t not in q.get("tema", "").lower():
            continue
        questoes.append(q)

    st.caption(f"Mostrando {len(questoes)} questoes")
    for q in questoes:
        with st.container(border=True):
            st.markdown(f"**📌 {q.get('tema', 'Sem tema')}** | Nivel: {q.get('dificuldade', '')} | ID: {q.get('id', '')}")
            perg = q.get("pergunta_direta", "")
            if len(perg) > 130:
                perg = perg[:130] + "..."
            st.write(perg)
            c1, c2 = st.columns([1, 1])
            if c1.button("Editar Rápido", key=f"cat_ed_{q['id']}"):
                dialog_editar_questao(q["id"])
            if c2.button("Excluir", key=f"cat_del_{q['id']}"):
                dialog_excluir_questao(q["id"])

# =====================================================================
# 19. TELA: GERAR PROVA (WORD)
# =====================================================================
def preparar_download_word(questoes, config, incluir_gabarito,
                           modo_acessibilidade, mostrar_descritor, titulo_prova):
    buffer = gerar_documento_word(
        questoes_selecionadas=questoes,
        config=config,
        incluir_gabarito=incluir_gabarito,
        modo_acessibilidade=modo_acessibilidade,
        mostrar_descritor=mostrar_descritor,
        titulo_prova=titulo_prova)
    st.session_state["prova_bytes"] = buffer.getvalue()
    st.session_state["prova_qtd"] = len(questoes)

def exibir_download_prova():
    if st.session_state.get("prova_bytes"):
        st.download_button(
            "Baixar Prova em Word (.docx)",
            data=st.session_state["prova_bytes"],
            file_name="Minha_Avaliacao.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", key="download_prova")
        st.caption(f"Prova gerada com {st.session_state.get('prova_qtd', 0)} questoes.")

def tela_gerar():
    st.markdown("## Gerar Avaliacao em Word")
    banco = carregar_banco()
    config = carregar_config()

    if not banco:
        st.error("O banco de questoes esta vazio!")
        return

    titulo_prova = st.text_input("Titulo da Avaliacao:",
                                 placeholder="Ex: AVALIACAO BIMESTRAL DE MATEMATICA")
    c_chk = st.columns(3)
    mostrar_descritor = c_chk[0].checkbox("Mostrar codigo do descritor (ex: D20)")
    incluir_gabarito = c_chk[1].checkbox("Incluir gabarito ao final", value=True)
    modo_acessibilidade = c_chk[2].checkbox("Modo Acessibilidade (Fonte 16pt, Coluna Unica)")

    st.markdown("---")
    st.markdown("### Opcao 1: Geracao Automatica por Conteudo")
    temas = sorted(set([q.get("tema", "Geral") for q in banco]))
    c1, c2 = st.columns([1, 2])
    qtd = int(c1.number_input("Qtd. de questoes:", min_value=1, max_value=len(banco),
                              value=min(10, len(banco)), step=1))
    temas_sel = c2.multiselect("Quais temas incluir? (vazio = todos)", temas)

    if st.button("Sortear e Gerar Prova", type="primary"):
        banco_filtrado = [q for q in banco if q.get("tema", "Geral") in temas_sel] if temas_sel else banco
        if len(banco_filtrado) == 0:
            st.error("Nenhuma questao encontrada para os temas selecionados.")
        elif qtd > len(banco_filtrado):
            st.error(f"Voce pediu {qtd} questoes, mas so existem {len(banco_filtrado)} disponiveis "
                     f"para os temas selecionados.")
        else:
            sorteadas = random.sample(banco_filtrado, qtd)
            preparar_download_word(sorteadas, config, incluir_gabarito,
                                   modo_acessibilidade, mostrar_descritor, titulo_prova)

    st.markdown("---")
    st.markdown("### Opcao 2: Selecao Manual (Avancada)")
    c1, c2 = st.columns(2)
    busca = c1.text_input("Buscar palavra no enunciado...", key="ger_busca")
    tema_f = c2.selectbox("Filtrar por tema:", ["Todos os Temas"] + temas, key="ger_tema")

    tb = busca.strip().lower()
    questoes_filtradas = []
    for q in banco:
        if tema_f != "Todos os Temas" and q.get("tema", "") != tema_f:
            continue
        if tb and tb not in q.get("enunciado", "").lower() \
                and tb not in q.get("pergunta_direta", "").lower():
            continue
        questoes_filtradas.append(q)

    selecionadas = []
    for q in questoes_filtradas:
        texto_chk = f"[{q.get('id')}] {q.get('tema', '')} - {q.get('pergunta_direta', '')[:110]}"
        if st.checkbox(texto_chk, key=f"sel_{q['id']}"):
            selecionadas.append(q)

    if st.button("Confirmar e Gerar Word", type="primary"):
        if not selecionadas:
            st.warning("Selecione pelo menos uma questao nas caixinhas!")
        else:
            preparar_download_word(selecionadas, config, incluir_gabarito,
                                   modo_acessibilidade, mostrar_descritor, titulo_prova)
    exibir_download_prova()

# =====================================================================
# 20. TELA: CONFIGURACOES
# =====================================================================
def tela_configuracoes():
    st.markdown("## Configuracoes do Sistema e Visual")
    config = carregar_config()

    with st.form("form_config"):
        st.markdown("**Dados institucionais:**")
        escola = st.text_input("Nome da Escola / Instituicao:", value=config.get("escola", ""))
        professor = st.text_input("Nome do Professor:", value=config.get("professor", ""))

        st.markdown("**Aparencia e Cores Customizadas:**")
        c1, c2 = st.columns(2)
        opcoes_ap = ["System", "Dark", "Light"]
        idx_ap = opcoes_ap.index(config.get("aparencia", "System")) \
            if config.get("aparencia", "System") in opcoes_ap else 0
        aparencia = c1.selectbox("Modo de Exibicao:", opcoes_ap, index=idx_ap)

        opcoes_tema = ["blue", "green", "dark-blue", "Personalizado"]
        idx_tema = opcoes_tema.index(config.get("cor_tema", "blue")) \
            if config.get("cor_tema", "blue") in opcoes_tema else 0
        cor_tema = c1.selectbox("Tema Padrao Base:", opcoes_tema, index=idx_tema)

        cor_principal = c1.color_picker("Cor Principal:", value=config.get("cor_principal", "#1f538d"))
        cor_secundaria = c1.color_picker("Cor Secundaria:", value=config.get("cor_secundaria", "#14375e"))
        cor_borda_card = c1.color_picker(
            "Cor da borda superior dos cards:",
            value=config.get("cor_borda_card") or config.get("cor_principal", "#1f538d"))
        c1.caption("As cores customizadas valem quando o 'Tema Padrao Base' for 'Personalizado'.")

        opcoes_fonte = ["Arial", "Times New Roman", "Calibri", "Tahoma"]
        idx_fonte = opcoes_fonte.index(config.get("fonte", "Arial")) \
            if config.get("fonte", "Arial") in opcoes_fonte else 0
        fonte = c2.selectbox("Fonte Padrao:", opcoes_fonte, index=idx_fonte)
        tamanho_fonte = c2.number_input("Tamanho da Fonte:", min_value=8.0, max_value=30.0,
                                        value=float(config.get("tamanho_fonte", 11)), step=0.5)
        margem_cm = c2.number_input("Margem (cm):", min_value=0.5, max_value=5.0,
                                    value=float(config.get("margem_cm", 1.5)), step=0.1)
        usar_duas_colunas = c2.checkbox("Usar 2 Colunas no Word",
                                        value=config.get("usar_duas_colunas", True))

        st.markdown("**Itens do Cabecalho das Provas:**")
        c3 = st.columns(3)
        mostrar_aluno = c3[0].checkbox("Mostrar linha de NOME DO ALUNO",
                                       value=config.get("mostrar_aluno", True))
        mostrar_turma = c3[1].checkbox("Mostrar linha de TURMA",
                                       value=config.get("mostrar_turma", True))
        mostrar_data = c3[2].checkbox("Mostrar linha de DATA",
                                      value=config.get("mostrar_data", True))

        salvar = st.form_submit_button("Salvar Configuracoes", type="primary",
                                       use_container_width=True)

    if salvar:
        try:
            nova_config = {
                "escola": escola.strip(),
                "professor": professor.strip(),
                "fonte": fonte,
                "tamanho_fonte": int(tamanho_fonte),
                "margem_cm": float(str(margem_cm).replace(",", ".")),
                "espacamento_pt": 0,
                "usar_duas_colunas": usar_duas_colunas,
                "mostrar_aluno": mostrar_aluno,
                "mostrar_turma": mostrar_turma,
                "mostrar_data": mostrar_data,
                "aparencia": aparencia,
                "cor_tema": cor_tema,
                "cor_principal": cor_principal.strip() or "#1f538d",
                "cor_secundaria": cor_secundaria.strip() or "#14375e",
                "cor_borda_card": cor_borda_card.strip() or config.get("cor_principal", "#1f538d")
            }
            salvar_config(nova_config)
            st.success("Configuracoes e cores aplicadas com sucesso!")
            st.rerun()
        except ValueError:
            st.error("Por favor, verifique se os campos numericos estao corretos.")

# =====================================================================
# 21. TELA DE BOAS-VINDAS (primeiro acesso)
# =====================================================================
def tela_onboarding():
    injetar_css(carregar_config())
    injetar_css_login()
    with st.container(key="login_card"):
        st.markdown(
            '<div class="login-logo">'
            f'<img class="login-img" src="data:image/png;base64,{LOGO_INICIO_B64}" '
            'alt="Exame Inteligente" /></div>'
            '<div class="login-titulo">Boas-vindas!<small>Complete seu perfil</small></div>'
            '<div class="login-sub">Sua conta foi criada. Conte para nos quem voce e '
            'para personalizar suas provas e relatorios.</div>',
            unsafe_allow_html=True)
        with st.form("form_onboarding"):
            ob_nome = st.text_input("Seu nome", key="ob_nome",
                                    placeholder="Como voce quer ser chamado")
            ob_escola = st.text_input("Nome da escola", key="ob_escola",
                                      placeholder="Ex.: Escola Municipal Paulo Freire")
            confirmar = st.form_submit_button("Comecar a usar", type="primary",
                                              use_container_width=True)
        if confirmar:
            if not ob_nome.strip():
                st.error("Informe o seu nome para continuar.")
            else:
                concluir_onboarding(ob_nome, ob_escola)
                st.rerun()

# =====================================================================
# 22. MAIN
# =====================================================================
def main():
    config = carregar_config()
    injetar_css(config)

    if not usuario_atual():
        if st.config.get_option("global.appTest") and os.environ.get("EXAME_TESTE_UI") != "1":
            st.session_state["usuario"] = garantir_usuario_teste()
        elif not _autologin_por_cookie():
            tela_login()
            injetar_js_movel()
            return

    usuario, token = garantir_cookie_token()
    injetar_js_movel(usuario, token)

    if st.session_state.get("bem_vindo"):
        st.session_state.pop("bem_vindo")

    if conta_precisa_onboarding():
        tela_onboarding()
        return

    montar_sidebar()
    pagina = pagina_atual()

    if pagina == "Dashboard":
        tela_dashboard(config)
    elif pagina == "Grade Semanal":
        tela_grade_semanal()
    elif pagina == "Turmas & Alunos":
        tela_turmas()
    elif pagina == "Central de Planos":
        tela_central_planos(config)
    elif pagina == "Anotacoes":
        tela_anotacoes()
    elif pagina == "Notas e Estatisticas":
        tela_notas()
    elif pagina == "Cadastrar Questao":
        tela_cadastrar()
    elif pagina == "Importar IA":
        tela_importar()
    elif pagina == "Catalogo de Questoes":
        tela_catalogo()
    elif pagina == "Gerar Prova":
        tela_gerar()
    elif pagina == "Configuracoes":
        tela_configuracoes()
    else:
        tela_dashboard(config)

if __name__ == "__main__":
    main()
