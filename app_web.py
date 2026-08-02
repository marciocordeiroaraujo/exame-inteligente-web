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
section[data-testid="stSidebar"] {color: #ececec !important;}
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown, section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
section[data-testid="stSidebar"] .stCaption {
    color: #ececec !important;
}
.logo-web {
    font-size: 2rem; font-weight: 800; line-height: 1.1; text-align: center;
    padding: 16px 10px 6px 10px; letter-spacing: -0.02em;
    background: linear-gradient(90deg, #ffffff, #e6c8ff);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
}
.logo-web small {font-size: .5em; letter-spacing: 4px; font-weight: 700;
    -webkit-text-fill-color: rgba(255,255,255,.70); color: rgba(255,255,255,.70);}
.nav-secao {
    font-size: .68rem; font-weight: 800; letter-spacing: 2.5px; color: rgba(255,255,255,.55) !important;
    padding: 18px 12px 6px 12px; text-transform: uppercase;
    border-top: 1px solid rgba(255,255,255,.08); margin-top: 8px;
}
section[data-testid="stSidebar"] .stButton > button {
    width: 100%; text-align: left; justify-content: flex-start; gap: 8px;
    background: transparent; border: 1px solid transparent; color: rgba(255,255,255,.85) !important;
    border-radius: 10px; padding: .58rem .9rem; font-size: .95rem; font-weight: 600;
    margin-bottom: 3px; transition: all .18s ease; box-shadow: none;
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
    box-shadow: 0 1px 3px rgba(0,0,0,.05); transition: box-shadow .2s ease, transform .2s ease;
}
.card:hover {box-shadow: 0 6px 20px rgba(0,0,0,.10); transform: translateY(-1px);}
/* Cards via st.container(key="card_...") -> classe st-key-card_* */
div[class*="st-key-card_"] {
    background: var(--card-bg); border: 1px solid var(--borda); border-radius: 14px;
    padding: 14px 16px; margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,.05); transition: box-shadow .2s ease;
}
div[class*="st-key-card_"]:hover {box-shadow: 0 6px 20px rgba(0,0,0,.10);}
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
    cor_fundo = config.get("cor_fundo", "#2b2b2b") if dark else "#f5f6f8"
    txt_btn = cor_texto_legivel(cor_p)
    card_bg = "#2b2b2b" if dark else "#ffffff"
    borda = "rgba(255,255,255,0.10)" if dark else "#e3e6ea"
    texto = "#e8e8e8" if dark else "#1f1f1f"
    texto_cinza = "#9a9a9a" if dark else "#6c757d"

    cor_p_hover = ajustar_cor(cor_p, 1.12)
    cor_p_deep = ajustar_cor(cor_p, 0.82)
    cor_p_edge = ajustar_cor(cor_p, 0.66)
    cor_s_grad = ajustar_cor(cor_s, 0.80)
    tag_txt = ajustar_cor(cor_s, 0.55)

    css = CSS_TEMPLATE
    css = css.replace("@@CORP@@", cor_p)
    css = css.replace("@@CORS@@", cor_s)
    css = css.replace("@@CORP_HOVER@@", cor_p_hover)
    css = css.replace("@@CORP_DEEP@@", cor_p_deep)
    css = css.replace("@@CORP_EDGE@@", cor_p_edge)
    css = css.replace("@@CORP_SOFT@@", cor_rgba(cor_p, 0.18))
    css = css.replace("@@CORP_SHADOW@@", cor_rgba(cor_p, 0.35))
    css = css.replace("@@CORS_GRAD@@", cor_s_grad)
    css = css.replace("@@CORS_SOFT@@", cor_rgba(cor_s, 0.12))
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
.login-logo { display: flex; align-items: center; gap: 14px; margin-bottom: 6px; }
.login-monogram {
    width: 54px; height: 54px; flex: 0 0 auto; border-radius: 16px;
    display: flex; align-items: center; justify-content: center;
    background: linear-gradient(135deg, #1f538d, #4f46e5);
    color: #ffffff !important; font-size: 1.3rem; font-weight: 800; letter-spacing: -.02em;
    box-shadow: 0 10px 22px rgba(31,83,141,.35);
}
.login-titulo {
    font-size: 1.45rem; font-weight: 800; color: #111827 !important;
    line-height: 1.15; letter-spacing: -.02em;
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

  // --- gesto: arrastar da borda esquerda abre seguindo o dedo; empurrar p/ esq fecha ---
  var EDGE = 30;
  var drag = null;

  // Bloqueia a navegacao de voltar por gesto do navegador (overscroll + touch-action)
  function injectStyles() {
    if (d.getElementById('ei-styles')) return;
    var s = d.createElement('style');
    s.id = 'ei-styles';
    s.textContent =
      'html, body { overscroll-behavior-x: contain !important; }' +
      '#ei-edge { position: fixed; left: 0; top: 0; bottom: 0; width: ' + EDGE + 'px;' +
      ' z-index: 2147483646; touch-action: pan-y; background: transparent; }';
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
    if (o) o.style.display = (isMobile() && !isOpen()) ? 'block' : 'none';
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

  function onStart(e) {
    if (!isMobile() || e.touches.length !== 1) return;
    var t = e.touches[0];
    drag = { sx: t.clientX, sy: t.clientY, dx: 0, dy: 0, mode: null };
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
      if (!open && drag.sx <= EDGE) drag.mode = 'open';
      else if (open && !inXScroll(e)) drag.mode = 'close';
      else { drag = null; return; }
    }
    e.preventDefault();
    var s = side(); if (!s) return;
    var w = sbWidth();
    s.style.transition = 'none';
    if (drag.mode === 'open') {
      // sidebar recolhida: revela seguindo o dedo (largura real + transform)
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
    var w = sbWidth();
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
      setTimeout(function() { if (isOpen()) closeSb(); }, 450);
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

  W.__ei = {
    isOpen: isOpen, openSb: openSb, closeSb: closeSb, toggleSb: toggleSb,
    sbWidth: sbWidth, isMobile: isMobile, syncCookie: syncCookie
  };

  var obs = new MutationObserver(function() { ensureFab(); watchNav(); ensureOverlay(); syncOverlay(); });
  if (d.body) obs.observe(d.body, { childList: true, subtree: true });
  setTimeout(function() { ensureOverlay(); syncOverlay(); }, 400);
  setInterval(function() { ensureOverlay(); syncOverlay(); }, 500);
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
    ("Dashboard", "Dashboard"),
    ("Grade Semanal", "Grade Semanal"),
    ("Turmas & Alunos", "Turmas & Alunos"),
    ("Central de Planos", "Central de Planos"),
    ("Anotacoes", "Anotacoes"),
]
NAV_AVALIACOES = [
    ("Notas e Estatisticas", "Notas e Estatisticas"),
    ("Cadastrar Questao", "Cadastrar Questao"),
    ("Importar IA", "Importar IA"),
    ("Catalogo de Questoes", "Catalogo de Questoes"),
    ("Gerar Prova", "Gerar Prova"),
    ("Configuracoes", "Configuracoes"),
]

def montar_sidebar():
    atual = pagina_atual()
    with st.sidebar:
        st.markdown(
            '<div class="logo-web">Exame<br>Inteligente<small>WEB</small></div>',
            unsafe_allow_html=True)
        st.markdown('<div class="nav-secao">Planejamento</div>', unsafe_allow_html=True)
        for chave, rotulo in NAV_PLANEJAMENTO:
            if st.button(rotulo, key=f"nav_p_{chave}",
                         type="primary" if atual == chave else "secondary",
                         use_container_width=True):
                st.query_params["pagina"] = chave
                st.rerun()
        st.markdown('<div class="nav-secao">Avaliacoes</div>', unsafe_allow_html=True)
        for chave, rotulo in NAV_AVALIACOES:
            if st.button(rotulo, key=f"nav_a_{chave}",
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
            '<div class="login-monogram">EI</div>'
            '<div class="login-titulo">Exame Inteligente<small>Plataforma WEB</small></div>'
            '</div>'
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

    col_esq, col_meio, col_dir = st.columns([1.7, 1, 1.1])

    with col_esq:
        with st.container(key="card_dash_cal"):
            novo_dia = render_calendario(planos)
            if novo_dia:
                st.session_state["dash_dia"] = novo_dia
                dia_selecionada = novo_dia

        with st.container(key="card_dash_grade"):
            st.markdown('<div class="card-titulo">Grade Semanal</div>', unsafe_allow_html=True)
            st.markdown(html_grade_mini(grade), unsafe_allow_html=True)

    with col_meio:
        with st.container(key="card_dash_mural"):
            topo = st.columns([3, 1])
            topo[0].markdown('<div class="card-titulo">Mural Rapido</div>', unsafe_allow_html=True)
            if topo[1].button("+ Novo", key="dash_novo_postit", use_container_width=True):
                dialog_postit()
            if not anotacoes:
                st.caption("Nenhum post-it. Crie um no botao + Novo.")
            for nota in anotacoes:
                cor = nota.get("cor") or "#fff3a3"
                txt = cor_texto_legivel(cor)
                st.markdown(
                    f'<div class="postit" style="background:{cor};color:{txt} !important;">'
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

    with col_dir:
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
    st.markdown("## Mural de Anotacoes (Post-its)")
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
            st.markdown(
                f'<div class="postit" style="background:{cor};min-height:120px;">'
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
                "cor_secundaria": cor_secundaria.strip() or "#14375e"
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
            '<div class="login-monogram">EI</div>'
            '<div class="login-titulo">Boas-vindas!<small>Complete seu perfil</small></div>'
            '</div>'
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
