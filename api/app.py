"""
ESTOQUE WPP — Backend unificado
Gerenciador de estoque, custos e vendas via WhatsApp (formulário ou IA/Groq)
"""
import os, json, csv, io, bcrypt, re, asyncio
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Optional

import psycopg2
import psycopg2.extras
from fastapi import FastAPI, HTTPException, Depends, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.responses import StreamingResponse, FileResponse
from pydantic import BaseModel
import jwt
import httpx

# ─────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────
SECRET_KEY    = os.getenv("SECRET_KEY", "troca-isso-em-producao")
ADMIN_TOKEN   = os.getenv("ADMIN_TOKEN", "troca-isso-tambem")
GROQ_API_KEY_1 = os.getenv("GROQ_API_KEY_1", os.getenv("GROQ_API_KEY", ""))
GROQ_API_KEY_2 = os.getenv("GROQ_API_KEY_2", "")
BAILEYS_URL   = os.getenv("BAILEYS_URL", "http://baileys:3000")
ALGORITHM     = "HS256"
TOKEN_EXP     = 24 * 7  # horas
DB_URL        = os.getenv("DATABASE_URL", "postgresql://postgres:postgres@postgres:5432/postgres")
GROQ_API_URL  = "https://api.groq.com/openai/v1/chat/completions"
TIMEZONE_PADRAO = ZoneInfo("America/Sao_Paulo")

GROQ_MODELOS_FALLBACK = [
    "llama-3.1-8b-instant",
    "llama-3.3-70b-versatile",
    "gemma2-9b-it",
]

# Ordem lógica: 1) operações do dia a dia, 2) relatórios, 3) cadastros,
# 4) automação/config, 5) ajuda. Os números aqui precisam bater com o
# roteamento em processar_texto() lá embaixo.
MENU_TEXTO = (
    "📋 *Menu*\n"
    "1️⃣ Registrar entrada de estoque\n"
    "2️⃣ Registrar venda\n"
    "3️⃣ Ajuste manual\n"
    "4️⃣ Consultar estoque de um produto\n"
    "5️⃣ Resumo do dia\n"
    "6️⃣ Visão geral do estoque\n"
    "7️⃣ Cadastrar produto\n"
    "8️⃣ Cadastrar matéria-prima\n"
    "9️⃣ Montar receita de um produto\n"
    "🔟 Configurar resumo automático\n"
    "1️⃣1️⃣ Ajuda — o que cada opção faz\n\n"
    "Responda com o número da opção."
)

TEXTO_AJUDA = (
    "ℹ️ *Como usar o sistema*\n\n"
    "1️⃣ *Entrada de estoque* — registra chegada de mercadoria (aumenta o estoque).\n"
    "2️⃣ *Venda* — registra uma venda (diminui o estoque e, se o produto tiver receita, "
    "desconta as matérias-primas usadas automaticamente).\n"
    "3️⃣ *Ajuste manual* — corrige o estoque de um produto pro valor exato que você digitar.\n"
    "4️⃣ *Consultar estoque* — mostra estoque, custo e preço de venda de um produto.\n"
    "5️⃣ *Resumo do dia* — total de entradas/vendas/ajustes de hoje.\n"
    "6️⃣ *Visão geral* — lista todos os produtos e matérias-primas com o estoque atual.\n"
    "7️⃣ *Cadastrar produto* — cria um novo produto (e opcionalmente já monta a receita dele).\n"
    "8️⃣ *Cadastrar matéria-prima* — cria um novo insumo usado nas receitas.\n"
    "9️⃣ *Montar receita* — define quais matérias-primas (e quantidades) um produto consome.\n"
    "🔟 *Resumo automático* — escolha até 2 horários por dia pra receber o resumo (opção 5) sem precisar pedir.\n\n"
    "A qualquer momento, digite *menu* para voltar aqui."
)

# ─────────────────────────────────────────
#  BANCO
# ─────────────────────────────────────────
def get_conn_raw():
    return psycopg2.connect(DB_URL, cursor_factory=psycopg2.extras.RealDictCursor)

def get_db():
    conn = get_conn_raw()
    try:
        yield conn
    finally:
        conn.close()

def db_one(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    row = cur.fetchone()
    return dict(row) if row else None

def db_all(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    return [dict(r) for r in cur.fetchall()]

def db_exec(conn, sql, params=()):
    cur = conn.cursor()
    cur.execute(sql, params)
    conn.commit()
    try:
        row = cur.fetchone()
        return dict(row) if row else None
    except Exception:
        return None

# ─────────────────────────────────────────
#  AUTH — clientes (painel)
# ─────────────────────────────────────────
security = HTTPBearer()

def criar_token(cliente_id: int, email: str) -> str:
    payload = {
        "sub": str(cliente_id),
        "email": email,
        "exp": datetime.utcnow() + timedelta(hours=TOKEN_EXP)
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)

def get_current_cliente(
    creds: HTTPAuthorizationCredentials = Depends(security),
    conn=Depends(get_db)
):
    try:
        payload = jwt.decode(creds.credentials, SECRET_KEY, algorithms=[ALGORITHM])
        cliente_id = int(payload["sub"])
    except Exception:
        raise HTTPException(status_code=401, detail="Token inválido")
    cliente = db_one(conn, "SELECT * FROM clientes WHERE id = %s AND ativo = TRUE", (cliente_id,))
    if not cliente:
        raise HTTPException(status_code=401, detail="Conta não encontrada ou inativa")
    return cliente

def check_admin(authorization: str = Header(default="")):
    """
    Aceita tanto 'Authorization: Bearer <token>' (usado pelo login.html)
    quanto o token puro no header, para compatibilidade.
    """
    token = authorization.replace("Bearer ", "").strip() if authorization else ""
    if not token or token != ADMIN_TOKEN:
        raise HTTPException(status_code=403, detail="Token de admin inválido")
    return True

# ─────────────────────────────────────────
#  SCHEMAS
# ─────────────────────────────────────────
class LoginBody(BaseModel):
    email: str
    senha: str

class AdminCriarClienteBody(BaseModel):
    nome_negocio: str
    email: str
    senha: str
    plano: str = "formulario"  # formulario | ia

class AdminPlanoBody(BaseModel):
    plano: str

class AdminAtivoBody(BaseModel):
    ativo: bool

class AdminNumeroBody(BaseModel):
    numero: str
    nome: Optional[str] = None

class AdminTrocaNumeroBody(BaseModel):
    numero_novo: str
    motivo: Optional[str] = None

class ProdutoBody(BaseModel):
    nome: str
    sku: Optional[str] = None
    custo_unitario: Optional[float] = 0
    preco_venda: Optional[float] = 0
    estoque_atual: Optional[float] = 0
    unidade: Optional[str] = "un"

class MovimentacaoManualBody(BaseModel):
    produto_id: int
    tipo: str  # entrada | saida | venda | ajuste
    quantidade: float
    valor_unitario: Optional[float] = 0

class MateriaPrimaBody(BaseModel):
    nome: str
    sku: Optional[str] = None
    custo_unitario: Optional[float] = 0
    estoque_atual: Optional[float] = 0
    unidade: Optional[str] = "un"
    estoque_minimo: Optional[float] = None  # opcional — se None, não gera alerta

class MovimentacaoMateriaPrimaBody(BaseModel):
    materia_prima_id: int
    tipo: str  # entrada | saida | ajuste
    quantidade: float
    valor_unitario: Optional[float] = 0

class ReceitaItemBody(BaseModel):
    materia_prima_id: int
    quantidade_necessaria: float

class ReceitaBody(BaseModel):
    itens: list[ReceitaItemBody]

# ─────────────────────────────────────────
#  HELPERS DE NEGÓCIO
# ─────────────────────────────────────────
def normalizar_numero(remote_jid_ou_numero: str) -> str:
    n = remote_jid_ou_numero.split("@")[0]
    return re.sub(r"\D", "", n)

def buscar_numero_autorizado(conn, numero: str):
    return db_one(conn, "SELECT * FROM numeros_autorizados WHERE numero = %s AND ativo = TRUE", (numero,))

def get_or_create_sessao(conn, numero_autorizado_id: int):
    sessao = db_one(conn, "SELECT * FROM sessoes_conversa WHERE numero_autorizado_id = %s", (numero_autorizado_id,))
    if sessao:
        return sessao
    return db_exec(conn, """
        INSERT INTO sessoes_conversa (numero_autorizado_id, etapa_atual, dados_parciais)
        VALUES (%s, 'menu', '{}') RETURNING *
    """, (numero_autorizado_id,))

def salvar_sessao(conn, numero_autorizado_id: int, etapa: str, dados: dict):
    db_exec(conn, """
        UPDATE sessoes_conversa
        SET etapa_atual = %s, dados_parciais = %s, atualizado_em = NOW()
        WHERE numero_autorizado_id = %s
    """, (etapa, json.dumps(dados), numero_autorizado_id))

def buscar_produto_por_nome(conn, cliente_id: int, nome: str):
    return db_one(conn, """
        SELECT * FROM produtos
        WHERE cliente_id = %s AND ativo = TRUE AND nome ILIKE %s
        ORDER BY id LIMIT 1
    """, (cliente_id, f"%{nome.strip()}%"))

def buscar_materia_prima_por_nome(conn, cliente_id: int, nome: str):
    return db_one(conn, """
        SELECT * FROM materias_primas
        WHERE cliente_id = %s AND ativo = TRUE AND nome ILIKE %s
        ORDER BY id LIMIT 1
    """, (cliente_id, f"%{nome.strip()}%"))

def fmt_num(valor) -> str:
    """Formata um número (Decimal ou float) sem casas decimais falsas
    (ex: evita mostrar '2.000' quando é só o número 2, o que em pt-BR
    é lido como 'dois mil'). Usa vírgula como separador decimal."""
    v = float(valor or 0)
    if v == int(v):
        texto = str(int(v))
    else:
        texto = f"{v:.3f}".rstrip("0").rstrip(".")
    return texto.replace(".", ",")

def listar_produtos_cliente(conn, cliente_id: int):
    return db_all(conn, """
        SELECT id, nome FROM produtos
        WHERE cliente_id = %s AND ativo = TRUE
        ORDER BY nome
    """, (cliente_id,))

def listar_materias_primas_cliente(conn, cliente_id: int):
    return db_all(conn, """
        SELECT id, nome, unidade FROM materias_primas
        WHERE cliente_id = %s AND ativo = TRUE
        ORDER BY nome
    """, (cliente_id,))

def montar_lista_numerada(itens, titulo: str, rodape: str = "Responda com o número.") -> str:
    linhas = [titulo, ""]
    for i, item in enumerate(itens, start=1):
        linhas.append(f"{i}. {item['nome']}")
    linhas.append("")
    linhas.append(rodape)
    return "\n".join(linhas)

def parse_selecao_multipla(texto: str, max_idx: int):
    """Aceita '1' ou '1,3,5' ou '1 3 5' e devolve a lista de índices (1-based)
    válidos, sem repetir, na ordem digitada. Retorna None se algo for inválido."""
    partes = re.split(r"[,\s]+", texto.strip())
    indices = []
    for p in partes:
        if not p:
            continue
        if not p.isdigit():
            return None
        idx = int(p)
        if idx < 1 or idx > max_idx:
            return None
        if idx not in indices:
            indices.append(idx)
    return indices or None

def avancar_fila_ou_confirmar(conn, numero_autorizado_id: int, dados: dict, tipo: str) -> str:
    """Chamado depois que um item do carrinho (produto + quantidade [+ valor])
    foi completado. Se ainda sobra produto na fila, pergunta a quantidade do
    próximo; se não sobra mais nada, monta o resumo do carrinho inteiro pra
    confirmação (SIM/NÃO)."""
    fila = dados.get("fila_produtos_ids", [])
    if fila:
        proximo_id = fila.pop(0)
        produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s", (proximo_id,))
        dados["fila_produtos_ids"] = fila
        dados["produto_id"] = produto["id"]
        dados["produto_nome"] = produto["nome"]
        salvar_sessao(conn, numero_autorizado_id, f"{tipo}_quantidade", dados)
        return f"Quantidade de *{produto['nome']}*?"

    salvar_sessao(conn, numero_autorizado_id, "confirmando", dados)
    carrinho = dados.get("carrinho", [])
    linhas = []
    total_geral = 0.0
    for item in carrinho:
        if tipo == "ajuste":
            linhas.append(f"- {item['produto_nome']} → {fmt_num(item['quantidade'])}")
        else:
            item_total = round(item["quantidade"] * item["valor_unitario"], 2)
            total_geral += item_total
            linhas.append(
                f"- {fmt_num(item['quantidade'])} × {item['produto_nome']} "
                f"a R$ {item['valor_unitario']:.2f} (R$ {item_total:.2f})"
            )
    corpo = "\n".join(linhas)
    acao = {"entrada": "a entrada", "venda": "a venda", "saida": "a saída", "ajuste": "os ajustes"}[tipo]
    if tipo == "ajuste":
        return f"Confirma {acao}?\n{corpo}\n\nResponda SIM ou NÃO."
    return f"Confirma {acao}?\n{corpo}\n\nTotal: R$ {total_geral:.2f}\n\nResponda SIM ou NÃO."

def aplicar_movimentacao(conn, cliente_id, produto_id, numero_autorizado_id, tipo, quantidade, valor_unitario, origem, mensagem_original=None):
    quantidade = float(quantidade)
    valor_unitario = float(valor_unitario or 0)
    valor_total = round(quantidade * valor_unitario, 2)

    db_exec(conn, """
        INSERT INTO movimentacoes
            (cliente_id, produto_id, numero_autorizado_id, tipo, quantidade, valor_unitario, valor_total, origem, mensagem_original)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, (cliente_id, produto_id, numero_autorizado_id, tipo, quantidade, valor_unitario, valor_total, origem, mensagem_original))

    alertas = []
    if tipo == "entrada":
        db_exec(conn, "UPDATE produtos SET estoque_atual = estoque_atual + %s, custo_unitario = %s WHERE id = %s",
                (quantidade, valor_unitario, produto_id))
    elif tipo in ("saida", "venda"):
        db_exec(conn, "UPDATE produtos SET estoque_atual = estoque_atual - %s WHERE id = %s", (quantidade, produto_id))
        if tipo == "venda" and valor_unitario:
            db_exec(conn, "UPDATE produtos SET preco_venda = %s WHERE id = %s", (valor_unitario, produto_id))
        if tipo == "venda":
            # Toda venda de um produto que tem receita cadastrada desconta
            # automaticamente a matéria-prima usada (ficha técnica / BOM).
            alertas = baixar_materia_prima_por_receita(conn, cliente_id, produto_id, numero_autorizado_id, quantidade, origem)
    elif tipo == "ajuste":
        db_exec(conn, "UPDATE produtos SET estoque_atual = %s WHERE id = %s", (quantidade, produto_id))

    return valor_total, alertas

# ─────────────────────────────────────────
#  MATÉRIA-PRIMA / RECEITA (ficha técnica)
# ─────────────────────────────────────────
def aplicar_movimentacao_materia_prima(conn, cliente_id, materia_prima_id, numero_autorizado_id, tipo,
                                        quantidade, valor_unitario, origem, mensagem_original=None,
                                        produto_id_origem=None):
    quantidade = float(quantidade)
    valor_unitario = float(valor_unitario or 0)
    valor_total = round(quantidade * valor_unitario, 2)

    db_exec(conn, """
        INSERT INTO movimentacoes_materia_prima
            (cliente_id, materia_prima_id, numero_autorizado_id, tipo, quantidade, valor_unitario,
             valor_total, origem, mensagem_original, produto_id_origem)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, (cliente_id, materia_prima_id, numero_autorizado_id, tipo, quantidade, valor_unitario,
          valor_total, origem, mensagem_original, produto_id_origem))

    if tipo == "entrada":
        db_exec(conn, "UPDATE materias_primas SET estoque_atual = estoque_atual + %s, custo_unitario = %s WHERE id = %s",
                (quantidade, valor_unitario, materia_prima_id))
    elif tipo in ("saida", "baixa_receita"):
        db_exec(conn, "UPDATE materias_primas SET estoque_atual = estoque_atual - %s WHERE id = %s",
                (quantidade, materia_prima_id))
    elif tipo == "ajuste":
        db_exec(conn, "UPDATE materias_primas SET estoque_atual = %s WHERE id = %s", (quantidade, materia_prima_id))

    # Alerta de estoque baixo — só dispara pra quem definiu um estoque mínimo
    # (campo opcional; quem não usa a função simplesmente nunca recebe isso).
    alerta = None
    if tipo in ("saida", "baixa_receita"):
        mp = db_one(conn, "SELECT nome, unidade, estoque_atual, estoque_minimo FROM materias_primas WHERE id = %s",
                    (materia_prima_id,))
        if mp and mp["estoque_minimo"] is not None and float(mp["estoque_atual"]) < float(mp["estoque_minimo"]):
            alerta = (f"⚠️ *{mp['nome']}* está com {fmt_num(mp['estoque_atual'])} {mp['unidade']}, "
                      f"abaixo do mínimo de {fmt_num(mp['estoque_minimo'])} {mp['unidade']}.")

    return valor_total, alerta

def baixar_materia_prima_por_receita(conn, cliente_id, produto_id, numero_autorizado_id, quantidade_vendida, origem):
    """Ao vender 1 ou mais unidades de um produto, desconta a matéria-prima
    de cada item da receita cadastrada, proporcionalmente à quantidade vendida.
    Produtos sem receita cadastrada simplesmente não têm nenhum item aqui.
    Retorna a lista de alertas de estoque baixo disparados (pode ser vazia)."""
    itens = db_all(conn, """
        SELECT r.materia_prima_id, r.quantidade_necessaria, m.custo_unitario
        FROM receita_itens r
        JOIN materias_primas m ON m.id = r.materia_prima_id
        WHERE r.produto_id = %s AND m.ativo = TRUE
    """, (produto_id,))
    alertas = []
    for item in itens:
        qtd_baixa = round(float(item["quantidade_necessaria"]) * float(quantidade_vendida), 4)
        _, alerta = aplicar_movimentacao_materia_prima(
            conn, cliente_id, item["materia_prima_id"], numero_autorizado_id,
            "baixa_receita", qtd_baixa, item["custo_unitario"],
            origem=origem, produto_id_origem=produto_id
        )
        if alerta:
            alertas.append(alerta)
    return alertas

# ─────────────────────────────────────────
#  IA (GROQ) — extração estruturada
# ─────────────────────────────────────────
def get_groq_key(cliente: dict) -> str:
    if cliente.get("groq_key_override"):
        return cliente["groq_key_override"]
    return GROQ_API_KEY_1 or GROQ_API_KEY_2

PROMPT_EXTRACAO = """Você é um extrator de dados de estoque. O usuário vai descrever uma movimentação em linguagem natural (entrada de mercadoria, venda ou saída).

Responda APENAS com um JSON válido, sem nenhum texto antes ou depois, no formato:
{"tipo": "entrada|venda|saida", "produto": "nome do produto", "quantidade": numero, "valor_unitario": numero}

Se não conseguir identificar algum campo com confiança, use null nesse campo.
Se a mensagem não for sobre estoque/venda, responda: {"tipo": null}
"""

async def chamar_groq_json(texto_usuario: str, groq_key: str) -> Optional[dict]:
    if not groq_key:
        return None
    messages = [
        {"role": "system", "content": PROMPT_EXTRACAO},
        {"role": "user", "content": texto_usuario}
    ]
    headers = {"Authorization": f"Bearer {groq_key}", "Content-Type": "application/json"}
    async with httpx.AsyncClient(timeout=25) as client:
        for modelo in GROQ_MODELOS_FALLBACK:
            payload = {"model": modelo, "temperature": 0.1, "max_tokens": 200, "messages": messages}
            try:
                resp = await client.post(GROQ_API_URL, headers=headers, json=payload)
                if resp.status_code != 200:
                    continue
                data = resp.json()
                bruto = data["choices"][0]["message"]["content"].strip()
                bruto = re.sub(r"^```json|```$", "", bruto, flags=re.MULTILINE).strip()
                return json.loads(bruto)
            except Exception:
                continue
    return None

async def enviar_whatsapp(destino: str, texto: str):
    # 'destino' pode ser um número puro (ex: "5511999998888") ou um JID completo
    # (ex: "224713024491669@lid" ou "5511999998888@s.whatsapp.net"). O baileys
    # só reconstrói "@s.whatsapp.net" quando não há "@" no valor — por isso é
    # essencial repassar o JID original (com @lid) quando ele existir, em vez
    # de normalizar para dígitos antes de responder.
    payload = {"number": destino, "message": texto}
    async with httpx.AsyncClient(timeout=30) as client:
        try:
            await client.post(f"{BAILEYS_URL}/disparar", json=payload)
        except Exception as e:
            print(f"⚠️ Erro ao enviar WhatsApp: {e}")

# ─────────────────────────────────────────
#  MÁQUINA DE ESTADOS — modo formulário
# ─────────────────────────────────────────
ETAPAS_TIPO = {"entrada": "entrada", "venda": "venda", "saida": "saida", "ajuste": "ajuste"}

def resposta_menu():
    return MENU_TEXTO

async def processar_texto(conn, cliente: dict, numero_autorizado: dict, texto: str) -> str:
    numero_autorizado_id = numero_autorizado["id"]
    sessao = get_or_create_sessao(conn, numero_autorizado_id)
    etapa = sessao["etapa_atual"]
    dados = sessao["dados_parciais"] if isinstance(sessao["dados_parciais"], dict) else json.loads(sessao["dados_parciais"] or "{}")
    texto_low = texto.strip().lower()

    if texto_low in ("menu", "cancelar", "0"):
        salvar_sessao(conn, numero_autorizado_id, "menu", {})
        return resposta_menu()

    # ── ETAPA: MENU ──
    if etapa == "menu":
        opcoes = {"1": "entrada", "2": "venda", "3": "ajuste", "4": "consulta"}
        if texto.strip() in opcoes:
            tipo = opcoes[texto.strip()]
            produtos = listar_produtos_cliente(conn, cliente["id"])
            if not produtos:
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return ("Você ainda não tem nenhum produto cadastrado. "
                        "Cadastre pelo painel (+ Novo Produto) e volte aqui depois.\n\n") + resposta_menu()
            dados = {"tipo": tipo, "produtos_ids": [p["id"] for p in produtos]}
            salvar_sessao(conn, numero_autorizado_id, f"{tipo}_produto", dados)
            rodape = "Responda com o número. Pra mais de um produto, separe por vírgula (ex: 1,3,5)."
            return montar_lista_numerada(produtos, "Qual produto?", rodape=rodape)

        if texto.strip() == "5":
            return await gerar_resumo_dia(conn, cliente["id"])

        if texto.strip() == "6":
            return gerar_visao_geral(conn, cliente["id"])

        if texto.strip() == "7":
            salvar_sessao(conn, numero_autorizado_id, "prod_nome", {})
            return "Qual o nome do novo produto?"

        if texto.strip() == "8":
            salvar_sessao(conn, numero_autorizado_id, "mp_nome", {})
            return "Qual o nome da nova matéria-prima?"

        if texto.strip() == "9":
            produtos = listar_produtos_cliente(conn, cliente["id"])
            if not produtos:
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return ("Você ainda não tem nenhum produto cadastrado. "
                        "Cadastre pelo painel (+ Novo Produto) e volte aqui depois.\n\n") + resposta_menu()
            dados = {"produtos_ids": [p["id"] for p in produtos]}
            salvar_sessao(conn, numero_autorizado_id, "receita_produto_escolha", dados)
            return montar_lista_numerada(produtos, "De qual produto você quer montar/editar a receita?")

        if texto.strip() == "10":
            config = obter_config_resumo_automatico(conn, cliente["id"])
            salvar_sessao(conn, numero_autorizado_id, "config_resumo_horarios", {})
            return (
                "🔟 *Resumo automático*\n\n"
                f"Horários atuais: {formatar_horarios_config(config)}\n\n"
                "Digite até 2 horários no formato HH:MM separados por vírgula "
                "(ex: 12:00,20:00) para receber o resumo do dia automaticamente nesses horários.\n"
                "Digite *desativar* para desligar o envio automático."
            )

        if texto.strip() == "11" or texto_low in ("ajuda", "help"):
            return TEXTO_AJUDA + "\n\n" + resposta_menu()

        # modo IA: tenta extrair da mensagem livre antes de cair no menu
        if cliente["plano"] == "ia":
            extraido = await chamar_groq_json(texto, get_groq_key(cliente))
            if extraido and extraido.get("tipo") in ("entrada", "venda", "saida"):
                return await preparar_confirmacao_ia(conn, cliente, numero_autorizado_id, extraido, texto)

        return "Não entendi 🤔\n\n" + resposta_menu()

    # ── ETAPA: escolhendo o(s) produto(s) da lista numerada ──
    if etapa.endswith("_produto"):
        tipo = dados.get("tipo")
        produtos_ids = dados.get("produtos_ids", [])
        indices = parse_selecao_multipla(texto, len(produtos_ids))
        if not indices:
            return (f"Manda o número do produto (1 a {len(produtos_ids)}). "
                     "Pra mais de um, separe por vírgula, ex: 1,3.")
        escolhidos_ids = [produtos_ids[i - 1] for i in indices]

        if tipo == "consulta":
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            blocos = []
            for pid in escolhidos_ids:
                produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s", (pid,))
                if not produto:
                    continue
                blocos.append(
                    f"📦 *{produto['nome']}*\n"
                    f"Estoque atual: {fmt_num(produto['estoque_atual'])} {produto['unidade']}\n"
                    f"Custo: R$ {produto['custo_unitario']:.2f}\n"
                    f"Preço de venda: R$ {produto['preco_venda']:.2f}"
                )
            return "\n\n".join(blocos) + "\n\n" + resposta_menu()

        # entrada / venda / ajuste: monta a fila e processa item por item
        primeiro = db_one(conn, "SELECT * FROM produtos WHERE id = %s", (escolhidos_ids[0],))
        dados["fila_produtos_ids"] = escolhidos_ids[1:]
        dados["carrinho"] = []
        dados["produto_id"] = primeiro["id"]
        dados["produto_nome"] = primeiro["nome"]
        salvar_sessao(conn, numero_autorizado_id, f"{tipo}_quantidade", dados)
        return f"Quantidade de *{primeiro['nome']}*?"

    # ── ETAPA: pedindo quantidade (de um item do carrinho) ──
    if etapa.endswith("_quantidade"):
        tipo = dados.get("tipo")
        try:
            quantidade = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número da quantidade, por favor."
        dados["quantidade"] = quantidade

        if tipo == "ajuste":
            dados.setdefault("carrinho", []).append({
                "produto_id": dados["produto_id"], "produto_nome": dados["produto_nome"],
                "quantidade": quantidade, "valor_unitario": 0,
            })
            return avancar_fila_ou_confirmar(conn, numero_autorizado_id, dados, tipo)

        # entrada / venda: sugere o valor já cadastrado no produto
        produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s", (dados["produto_id"],))
        valor_sugerido = float((produto["custo_unitario"] if tipo == "entrada" else produto["preco_venda"]) or 0)
        dados["valor_sugerido"] = valor_sugerido
        salvar_sessao(conn, numero_autorizado_id, f"{tipo}_valor", dados)
        rotulo = "custo unitário" if tipo == "entrada" else "valor unitário de venda"
        if valor_sugerido > 0:
            return (
                f"{rotulo.capitalize()} de *{dados['produto_nome']}*: R$ {valor_sugerido:.2f} (já cadastrado no painel).\n"
                f"Responda OK para manter esse valor, ou digite outro valor."
            )
        return f"Qual o {rotulo} (R$) de *{dados['produto_nome']}*?"

    # ── ETAPA: pedindo/confirmando valor de um item do carrinho ──
    if etapa.endswith("_valor"):
        tipo = dados.get("tipo")
        valor_sugerido = dados.get("valor_sugerido", 0) or 0
        if texto_low in ("ok", "sim", "s", "manter", "confirmo") and valor_sugerido > 0:
            valor = valor_sugerido
        else:
            try:
                valor = float(texto.replace(",", "."))
            except ValueError:
                if valor_sugerido > 0:
                    return "Responda OK para manter o valor sugerido, ou digite um número."
                return "Manda o valor em número, por favor."

        dados.setdefault("carrinho", []).append({
            "produto_id": dados["produto_id"], "produto_nome": dados["produto_nome"],
            "quantidade": dados["quantidade"], "valor_unitario": valor,
        })
        return avancar_fila_ou_confirmar(conn, numero_autorizado_id, dados, tipo)

    # ── ETAPA: cadastro de matéria-prima (opção 6) ──
    if etapa == "mp_nome":
        dados["mp_nome"] = texto.strip()
        salvar_sessao(conn, numero_autorizado_id, "mp_unidade", dados)
        return "Qual a unidade de medida? (ex: kg, g, l, ml, un) — ou digite PULAR para usar 'un'"

    if etapa == "mp_unidade":
        dados["mp_unidade"] = texto.strip() if texto_low != "pular" else "un"
        salvar_sessao(conn, numero_autorizado_id, "mp_custo", dados)
        return "Qual o custo unitário (R$)? — digite 0 se ainda não souber"

    if etapa == "mp_custo":
        try:
            custo = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número do custo, por favor."
        dados["mp_custo"] = custo
        salvar_sessao(conn, numero_autorizado_id, "mp_estoque", dados)
        return "Qual o estoque inicial dessa matéria-prima?"

    if etapa == "mp_estoque":
        try:
            estoque = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número do estoque, por favor."
        dados["mp_estoque"] = estoque
        salvar_sessao(conn, numero_autorizado_id, "mp_estoque_minimo", dados)
        return (
            "Quer receber um aviso quando essa matéria-prima ficar baixa? "
            "Se sim, digite o estoque mínimo (ex: 1). Se não quiser usar isso, digite PULAR."
        )

    if etapa == "mp_estoque_minimo":
        if texto_low == "pular":
            dados["mp_estoque_minimo"] = None
        else:
            try:
                dados["mp_estoque_minimo"] = float(texto.replace(",", "."))
            except ValueError:
                return "Manda só o número do estoque mínimo, ou digite PULAR pra não usar essa opção."
        salvar_sessao(conn, numero_autorizado_id, "confirmando_mp", dados)
        linha_minimo = (f" | alerta abaixo de {fmt_num(dados['mp_estoque_minimo'])}"
                         if dados.get("mp_estoque_minimo") is not None else "")
        return (
            f"Confirma o cadastro?\n"
            f"*{dados['mp_nome']}* | {dados['mp_unidade']} | custo R$ {dados['mp_custo']:.2f} | "
            f"estoque inicial {fmt_num(dados['mp_estoque'])}{linha_minimo}\n\nResponda SIM ou NÃO."
        )

    if etapa == "confirmando_mp":
        if texto_low in ("sim", "s", "confirmo", "confirmar"):
            existente = buscar_materia_prima_por_nome(conn, cliente["id"], dados["mp_nome"])
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            if existente:
                return f"⚠️ Já existe uma matéria-prima chamada '{existente['nome']}'. Use o painel pra editar.\n\n" + resposta_menu()
            db_exec(conn, """
                INSERT INTO materias_primas (cliente_id, nome, unidade, custo_unitario, estoque_atual, estoque_minimo)
                VALUES (%s,%s,%s,%s,%s,%s)
            """, (cliente["id"], dados["mp_nome"], dados["mp_unidade"], dados["mp_custo"], dados["mp_estoque"],
                  dados.get("mp_estoque_minimo")))
            return "✅ Matéria-prima cadastrada com sucesso!\n\n" + resposta_menu()
        if texto_low in ("não", "nao", "n"):
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Cancelado.\n\n" + resposta_menu()
        return "Responda SIM ou NÃO."

    # ── ETAPA: cadastro de produto (opção 8) ──
    if etapa == "prod_nome":
        dados["prod_nome"] = texto.strip()
        salvar_sessao(conn, numero_autorizado_id, "prod_unidade", dados)
        return "Qual a unidade de medida? (ex: un, kg, l) — ou digite PULAR para usar 'un'"

    if etapa == "prod_unidade":
        dados["prod_unidade"] = texto.strip() if texto_low != "pular" else "un"
        salvar_sessao(conn, numero_autorizado_id, "prod_custo", dados)
        return "Qual o custo unitário (R$)? — digite 0 se ainda não souber"

    if etapa == "prod_custo":
        try:
            custo = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número do custo, por favor."
        dados["prod_custo"] = custo
        salvar_sessao(conn, numero_autorizado_id, "prod_preco", dados)
        return "Qual o valor de venda (R$)? — digite 0 se ainda não souber"

    if etapa == "prod_preco":
        try:
            preco = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número do valor de venda, por favor."
        dados["prod_preco"] = preco
        salvar_sessao(conn, numero_autorizado_id, "prod_estoque", dados)
        return "Qual o estoque inicial desse produto?"

    if etapa == "prod_estoque":
        try:
            estoque = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número do estoque, por favor."
        dados["prod_estoque"] = estoque
        salvar_sessao(conn, numero_autorizado_id, "prod_quer_receita", dados)
        return (
            "Esse produto usa alguma matéria-prima na receita (ex: farinha, embalagem)? "
            "Se sim, a matéria-prima é descontada sozinha a cada venda. Responda SIM ou NÃO."
        )

    if etapa == "prod_quer_receita":
        if texto_low in ("sim", "s"):
            dados["prod_quer_receita"] = True
        elif texto_low in ("não", "nao", "n"):
            dados["prod_quer_receita"] = False
        else:
            return "Responda SIM ou NÃO."
        salvar_sessao(conn, numero_autorizado_id, "produto_aguardando_confirmacao", dados)
        return (
            f"Confirma o cadastro?\n"
            f"*{dados['prod_nome']}* | {dados['prod_unidade']} | custo R$ {dados['prod_custo']:.2f} | "
            f"venda R$ {dados['prod_preco']:.2f} | estoque inicial {fmt_num(dados['prod_estoque'])}"
            f"\n\nResponda SIM ou NÃO."
        )

    if etapa == "produto_aguardando_confirmacao":
        if texto_low in ("sim", "s", "confirmo", "confirmar"):
            existente = buscar_produto_por_nome(conn, cliente["id"], dados["prod_nome"])
            if existente:
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return f"⚠️ Já existe um produto chamado '{existente['nome']}'. Use o painel pra editar.\n\n" + resposta_menu()
            produto_novo = db_exec(conn, """
                INSERT INTO produtos (cliente_id, nome, unidade, custo_unitario, preco_venda, estoque_atual)
                VALUES (%s,%s,%s,%s,%s,%s) RETURNING *
            """, (cliente["id"], dados["prod_nome"], dados["prod_unidade"], dados["prod_custo"],
                  dados["prod_preco"], dados["prod_estoque"]))

            if not dados.get("prod_quer_receita"):
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return "✅ Produto cadastrado com sucesso!\n\n" + resposta_menu()

            # Quis vincular receita — encadeia direto no mesmo fluxo da opção 7,
            # já com o produto recém-criado.
            materias = listar_materias_primas_cliente(conn, cliente["id"])
            if not materias:
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return ("✅ Produto cadastrado com sucesso!\n\n"
                        "Só que você ainda não tem matéria-prima cadastrada. Cadastre uma (opção 6) "
                        "e depois monte a receita pela opção 7.\n\n") + resposta_menu()

            dados_receita = {
                "receita_produto_id": produto_novo["id"], "receita_produto_nome": produto_novo["nome"],
                "receita_itens": [], "materias_ids": [m["id"] for m in materias],
            }
            salvar_sessao(conn, numero_autorizado_id, "receita_item_escolha", dados_receita)
            return "✅ Produto cadastrado! Agora vamos montar a receita.\n\n" + montar_lista_numerada(
                materias, f"Qual matéria-prima entra em *{produto_novo['nome']}*?",
                rodape="Responda com o número, ou digite PRONTO quando terminar."
            )
        if texto_low in ("não", "nao", "n"):
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Cancelado.\n\n" + resposta_menu()
        return "Responda SIM ou NÃO."

    # ── ETAPA: montar receita de um produto (opção 7) ──
    if etapa == "receita_produto_escolha":
        produtos_ids = dados.get("produtos_ids", [])
        try:
            idx = int(texto.strip())
            assert 1 <= idx <= len(produtos_ids)
        except (ValueError, AssertionError):
            return f"Manda só o número do produto (1 a {len(produtos_ids)})."
        produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s", (produtos_ids[idx - 1],))
        if not produto:
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Esse produto não existe mais.\n\n" + resposta_menu()

        materias = listar_materias_primas_cliente(conn, cliente["id"])
        if not materias:
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Você ainda não tem matéria-prima cadastrada. Cadastre uma primeiro (opção 6).\n\n" + resposta_menu()

        dados = {
            "receita_produto_id": produto["id"], "receita_produto_nome": produto["nome"],
            "receita_itens": [], "materias_ids": [m["id"] for m in materias],
        }
        salvar_sessao(conn, numero_autorizado_id, "receita_item_escolha", dados)
        return montar_lista_numerada(
            materias, f"Montando a receita de *{produto['nome']}*.\nQual matéria-prima entra nela?",
            rodape="Responda com o número, ou digite PRONTO quando terminar."
        )

    if etapa == "receita_item_escolha":
        if texto_low == "pronto":
            if not dados.get("receita_itens"):
                salvar_sessao(conn, numero_autorizado_id, "menu", {})
                return "Nenhum item adicionado. Receita cancelada.\n\n" + resposta_menu()
            salvar_sessao(conn, numero_autorizado_id, "confirmando_receita", dados)
            linhas = "\n".join(f"- {fmt_num(i['quantidade'])} {i['unidade']} de {i['nome']}" for i in dados["receita_itens"])
            return f"Confirma a receita de *{dados['receita_produto_nome']}*?\n{linhas}\n\nResponda SIM ou NÃO."

        materias_ids = dados.get("materias_ids", [])
        try:
            idx = int(texto.strip())
            assert 1 <= idx <= len(materias_ids)
        except (ValueError, AssertionError):
            return f"Manda só o número da matéria-prima (1 a {len(materias_ids)}), ou PRONTO para terminar."
        materia = db_one(conn, "SELECT * FROM materias_primas WHERE id = %s", (materias_ids[idx - 1],))
        if not materia:
            return "Essa matéria-prima não existe mais. Escolha outro número ou digite PRONTO."
        dados["receita_item_atual"] = {"id": materia["id"], "nome": materia["nome"], "unidade": materia["unidade"]}
        salvar_sessao(conn, numero_autorizado_id, "receita_item_qtd", dados)
        return f"Quantos {materia['unidade']} de *{materia['nome']}* vão em 1 unidade do produto?"

    if etapa == "receita_item_qtd":
        try:
            quantidade = float(texto.replace(",", "."))
        except ValueError:
            return "Manda só o número da quantidade, por favor."
        item = dados["receita_item_atual"]
        dados.setdefault("receita_itens", []).append({
            "materia_prima_id": item["id"], "nome": item["nome"], "unidade": item["unidade"], "quantidade": quantidade
        })
        dados.pop("receita_item_atual", None)
        salvar_sessao(conn, numero_autorizado_id, "receita_item_escolha", dados)
        materias = listar_materias_primas_cliente(conn, cliente["id"])
        dados["materias_ids"] = [m["id"] for m in materias]
        salvar_sessao(conn, numero_autorizado_id, "receita_item_escolha", dados)
        return montar_lista_numerada(
            materias, "Adicionado! Mais alguma matéria-prima?",
            rodape="Responda com o número, ou digite PRONTO para terminar."
        )

    if etapa == "confirmando_receita":
        if texto_low in ("sim", "s", "confirmo", "confirmar"):
            produto_id = dados["receita_produto_id"]
            db_exec(conn, "DELETE FROM receita_itens WHERE produto_id = %s", (produto_id,))
            for item in dados["receita_itens"]:
                db_exec(conn, """
                    INSERT INTO receita_itens (produto_id, materia_prima_id, quantidade_necessaria)
                    VALUES (%s,%s,%s)
                """, (produto_id, item["materia_prima_id"], item["quantidade"]))
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "✅ Receita salva com sucesso! A partir de agora, vender esse produto já desconta a matéria-prima automaticamente.\n\n" + resposta_menu()
        if texto_low in ("não", "nao", "n"):
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Cancelado.\n\n" + resposta_menu()
        return "Responda SIM ou NÃO."

    # ── ETAPA: confirmando ──
    if etapa == "confirmando":
        if texto_low in ("sim", "s", "confirmo", "confirmar"):
            tipo = dados.get("tipo")
            carrinho = dados.get("carrinho")
            todos_alertas = []
            if carrinho:
                for item in carrinho:
                    _, alertas_item = aplicar_movimentacao(
                        conn, cliente["id"], item["produto_id"], numero_autorizado_id,
                        tipo, item["quantidade"], item.get("valor_unitario", 0),
                        origem="formulario", mensagem_original=texto
                    )
                    todos_alertas.extend(alertas_item)
            else:
                # compatibilidade com o modo IA, que ainda manda um único item
                _, alertas_item = aplicar_movimentacao(
                    conn, cliente["id"], dados["produto_id"], numero_autorizado_id,
                    tipo, dados["quantidade"], dados.get("valor_unitario", 0),
                    origem="formulario", mensagem_original=texto
                )
                todos_alertas.extend(alertas_item)
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            resposta = "✅ Registrado com sucesso!"
            if todos_alertas:
                resposta += "\n\n" + "\n".join(todos_alertas)
            return resposta + "\n\n" + resposta_menu()
        if texto_low in ("não", "nao", "n"):
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "Cancelado.\n\n" + resposta_menu()
        return "Responda SIM ou NÃO."

    # ── ETAPA: configurando horário(s) do resumo automático (opção 10) ──
    if etapa == "config_resumo_horarios":
        if texto_low in ("desativar", "desligar", "remover", "cancelar_config"):
            salvar_config_resumo_automatico(conn, cliente["id"], None, None, ativo=False)
            salvar_sessao(conn, numero_autorizado_id, "menu", {})
            return "🔕 Resumo automático desativado.\n\n" + resposta_menu()

        horarios = parse_horarios(texto)
        if horarios is None:
            return (
                "Não entendi os horários 🤔\n"
                "Digite até 2 horários no formato HH:MM separados por vírgula (ex: 12:00,20:00), "
                "ou *desativar* para desligar."
            )
        horario_1 = horarios[0]
        horario_2 = horarios[1] if len(horarios) > 1 else None
        salvar_config_resumo_automatico(conn, cliente["id"], horario_1, horario_2, ativo=True)
        salvar_sessao(conn, numero_autorizado_id, "menu", {})
        resumo_horarios = " e ".join(horarios)
        return (
            f"✅ Resumo automático configurado para {resumo_horarios} (todos os números autorizados recebem).\n\n"
            + resposta_menu()
        )

    # fallback de segurança
    salvar_sessao(conn, numero_autorizado_id, "menu", {})
    return resposta_menu()

async def preparar_confirmacao_ia(conn, cliente, numero_autorizado_id, extraido, texto_original):
    tipo = extraido.get("tipo")
    nome_produto = extraido.get("produto")
    quantidade = extraido.get("quantidade")
    valor_unitario = extraido.get("valor_unitario") or 0

    if not nome_produto or quantidade is None:
        return "Entendi que é sobre estoque, mas faltou produto ou quantidade. Pode reescrever?\n\n" + resposta_menu()

    produto = buscar_produto_por_nome(conn, cliente["id"], nome_produto)
    if not produto:
        produto = db_exec(conn, "INSERT INTO produtos (cliente_id, nome) VALUES (%s,%s) RETURNING *",
                           (cliente["id"], nome_produto.strip()))

    dados = {
        "tipo": tipo, "produto_id": produto["id"], "produto_nome": produto["nome"],
        "quantidade": float(quantidade), "valor_unitario": float(valor_unitario or 0),
    }
    salvar_sessao(conn, numero_autorizado_id, "confirmando", dados)
    total = round(dados["quantidade"] * dados["valor_unitario"], 2)
    acao = {"entrada": "Entrada de", "venda": "Venda de", "saida": "Saída de"}[tipo]
    return (
        f"Confirma?\n{acao} {fmt_num(dados['quantidade'])} × *{produto['nome']}* "
        f"a R$ {dados['valor_unitario']:.2f} (total R$ {total:.2f})\n\nResponda SIM ou NÃO."
    )

def gerar_visao_geral(conn, cliente_id: int) -> str:
    """Lista todos os produtos e matérias-primas com o estoque atual,
    sinalizando com ⚠️ quem tem estoque mínimo definido e já ficou abaixo dele."""
    produtos = db_all(conn, """
        SELECT nome, unidade, estoque_atual FROM produtos
        WHERE cliente_id = %s AND ativo = TRUE ORDER BY nome
    """, (cliente_id,))
    materias = db_all(conn, """
        SELECT nome, unidade, estoque_atual, estoque_minimo FROM materias_primas
        WHERE cliente_id = %s AND ativo = TRUE ORDER BY nome
    """, (cliente_id,))

    blocos = ["📊 *Visão geral do estoque*"]

    blocos.append("\n📦 *Produtos*")
    if produtos:
        for p in produtos:
            blocos.append(f"- {p['nome']}: {fmt_num(p['estoque_atual'])} {p['unidade']}")
    else:
        blocos.append("Nenhum produto cadastrado ainda.")

    blocos.append("\n🧂 *Matérias-primas*")
    if materias:
        for m in materias:
            abaixo = m["estoque_minimo"] is not None and float(m["estoque_atual"]) < float(m["estoque_minimo"])
            marca = "⚠️ " if abaixo else "- "
            blocos.append(f"{marca}{m['nome']}: {fmt_num(m['estoque_atual'])} {m['unidade']}")
    else:
        blocos.append("Nenhuma matéria-prima cadastrada ainda.")

    return "\n".join(blocos) + "\n\n" + resposta_menu()

async def gerar_resumo_dia(conn, cliente_id: int) -> str:
    hoje = datetime.utcnow().date()
    linhas = db_all(conn, """
        SELECT tipo, COALESCE(SUM(quantidade),0) qtd, COALESCE(SUM(valor_total),0) total
        FROM movimentacoes
        WHERE cliente_id = %s AND criado_em::date = %s
        GROUP BY tipo
    """, (cliente_id, hoje))
    if not linhas:
        return f"📊 Nenhuma movimentação hoje ({hoje.strftime('%d/%m')}).\n\n" + resposta_menu()
    txt = f"📊 *Resumo de hoje ({hoje.strftime('%d/%m')})*\n"
    for l in linhas:
        txt += f"- {l['tipo'].capitalize()}: {fmt_num(l['qtd'])} un | R$ {float(l['total']):.2f}\n"
    return txt + "\n" + resposta_menu()

# ─────────────────────────────────────────
#  RESUMO AUTOMÁTICO — configuração e disparo agendado
# ─────────────────────────────────────────
def parse_horarios(texto: str):
    """Aceita até 2 horários no formato HH:MM separados por vírgula/espaço
    (ex: '12:00,20:00' ou '12:00 20:00'). Retorna lista de strings 'HH:MM'
    normalizadas, ou None se algo for inválido."""
    partes = [p.strip() for p in re.split(r"[,\s]+", texto.strip()) if p.strip()]
    if not partes or len(partes) > 2:
        return None
    horarios = []
    for p in partes:
        m = re.fullmatch(r"(\d{1,2}):(\d{2})", p)
        if not m:
            return None
        h, mi = int(m.group(1)), int(m.group(2))
        if h > 23 or mi > 59:
            return None
        hhmm = f"{h:02d}:{mi:02d}"
        if hhmm not in horarios:
            horarios.append(hhmm)
    return horarios or None

def obter_config_resumo_automatico(conn, cliente_id: int):
    return db_one(conn, "SELECT * FROM resumo_automatico_config WHERE cliente_id = %s", (cliente_id,))

def formatar_horarios_config(config) -> str:
    if not config or not config.get("ativo") or not (config.get("horario_1") or config.get("horario_2")):
        return "nenhum configurado"
    horarios = [str(h)[:5] for h in (config.get("horario_1"), config.get("horario_2")) if h]
    return " e ".join(horarios)

def salvar_config_resumo_automatico(conn, cliente_id: int, horario_1, horario_2, ativo: bool):
    db_exec(conn, """
        INSERT INTO resumo_automatico_config (cliente_id, horario_1, horario_2, ativo, atualizado_em)
        VALUES (%s, %s, %s, %s, NOW())
        ON CONFLICT (cliente_id) DO UPDATE
        SET horario_1 = EXCLUDED.horario_1,
            horario_2 = EXCLUDED.horario_2,
            ativo = EXCLUDED.ativo,
            atualizado_em = NOW()
    """, (cliente_id, horario_1, horario_2, ativo))

async def disparar_resumo_automatico_cliente(conn, cliente_id: int):
    resumo = await gerar_resumo_dia(conn, cliente_id)
    numeros = db_all(conn, "SELECT numero FROM numeros_autorizados WHERE cliente_id = %s AND ativo = TRUE",
                      (cliente_id,))
    for n in numeros:
        await enviar_whatsapp(n["numero"], "⏰ *Resumo automático*\n\n" + resumo)

async def checar_e_disparar_resumos_automaticos():
    """Roda 1x por minuto. Compara o horário atual (America/Sao_Paulo) com os
    horários configurados por cada cliente e dispara o resumo do dia quando
    bate, evitando reenvio duplicado no mesmo dia via ultimo_envio_1/2."""
    agora = datetime.now(TIMEZONE_PADRAO)
    hoje = agora.date()
    hora_atual = agora.strftime("%H:%M")

    conn = get_conn_raw()
    try:
        configs = db_all(conn, "SELECT * FROM resumo_automatico_config WHERE ativo = TRUE")
        for c in configs:
            for campo_horario, campo_envio in (("horario_1", "ultimo_envio_1"), ("horario_2", "ultimo_envio_2")):
                horario = c.get(campo_horario)
                if not horario:
                    continue
                if str(horario)[:5] != hora_atual:
                    continue
                if c.get(campo_envio) == hoje:
                    continue  # já disparou esse horário hoje
                try:
                    await disparar_resumo_automatico_cliente(conn, c["cliente_id"])
                except Exception as e:
                    print(f"⚠️ Erro ao disparar resumo automático (cliente {c['cliente_id']}): {e}")
                db_exec(conn, f"UPDATE resumo_automatico_config SET {campo_envio} = %s WHERE cliente_id = %s",
                        (hoje, c["cliente_id"]))
    finally:
        conn.close()

async def loop_relogio_resumo_automatico():
    while True:
        try:
            await checar_e_disparar_resumos_automaticos()
        except Exception as e:
            print(f"⚠️ Erro no relógio de resumo automático: {e}")
        await asyncio.sleep(60)

# ─────────────────────────────────────────
#  STARTUP — Criar tabelas se não existirem
# ─────────────────────────────────────────
def criar_tabelas_resumo_automatico():
    """Cria a tabela resumo_automatico_config no startup SE ela não existir.
    Seguro: usa IF NOT EXISTS, nunca deleta ou altera dados existentes."""
    conn = get_conn_raw()
    try:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS resumo_automatico_config (
                cliente_id      INTEGER PRIMARY KEY REFERENCES clientes(id) ON DELETE CASCADE,
                horario_1       TIME NULL,
                horario_2       TIME NULL,
                ativo           BOOLEAN NOT NULL DEFAULT TRUE,
                ultimo_envio_1  DATE NULL,
                ultimo_envio_2  DATE NULL,
                atualizado_em   TIMESTAMP NOT NULL DEFAULT NOW()
            );
        """)
        conn.commit()
        print("✅ Tabela resumo_automatico_config: OK")
    except Exception as e:
        print(f"⚠️ Erro ao criar tabela: {e}")
    finally:
        conn.close()

# ─────────────────────────────────────────
#  LIFESPAN
# ─────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    # 1️⃣ Criar tabelas necessárias
    criar_tabelas_resumo_automatico()
    
    # 2️⃣ Testar conexão com banco
    try:
        conn = get_conn_raw()
        conn.close()
        print("✅ Conexão com banco OK")
    except Exception as e:
        print(f"⚠️ Não foi possível conectar ao banco no startup: {e}")

    # 3️⃣ Iniciar o "relógio" de resumo automático
    tarefa_relogio = asyncio.create_task(loop_relogio_resumo_automatico())
    yield
    
    # 4️⃣ Cleanup: cancelar o relógio quando o app fecha
    tarefa_relogio.cancel()
    try:
        await tarefa_relogio
    except asyncio.CancelledError:
        pass

app = FastAPI(title="Estoque WPP", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"],
)

# ═════════════════════════════════════════
#  AUTH — CLIENTE
# ═════════════════════════════════════════
@app.post("/auth/login")
def login(body: LoginBody, conn=Depends(get_db)):
    cliente = db_one(conn, "SELECT * FROM clientes WHERE email = %s AND ativo = TRUE", (body.email,))
    if not cliente or not bcrypt.checkpw(body.senha.encode(), cliente["senha_hash"].encode()):
        raise HTTPException(status_code=401, detail="Email ou senha inválidos")
    token = criar_token(cliente["id"], cliente["email"])
    return {
        "access_token": token,
        "cliente_id": cliente["id"],
        "cliente": {k: v for k, v in cliente.items() if k != "senha_hash"},
    }

@app.get("/auth/me")
def me(cliente=Depends(get_current_cliente)):
    return {k: v for k, v in cliente.items() if k != "senha_hash"}

# ═════════════════════════════════════════
#  ADMIN — gestão de clientes/números/conexão
# ═════════════════════════════════════════
@app.get("/admin/clientes")
def admin_listar_clientes(conn=Depends(get_db), _admin=Depends(check_admin)):
    clientes = db_all(conn, "SELECT id, nome_negocio, email, plano, ativo, criado_em FROM clientes ORDER BY id DESC")
    for c in clientes:
        c["numeros"] = db_all(conn, "SELECT * FROM numeros_autorizados WHERE cliente_id = %s", (c["id"],))
    return clientes

@app.post("/admin/clientes")
def admin_criar_cliente(body: AdminCriarClienteBody, conn=Depends(get_db), _admin=Depends(check_admin)):
    if db_one(conn, "SELECT id FROM clientes WHERE email = %s", (body.email,)):
        raise HTTPException(status_code=400, detail="Email já cadastrado")
    senha_hash = bcrypt.hashpw(body.senha.encode(), bcrypt.gensalt()).decode()
    cliente = db_exec(conn, """
        INSERT INTO clientes (nome_negocio, email, senha_hash, plano)
        VALUES (%s,%s,%s,%s) RETURNING id, nome_negocio, email, plano, ativo
    """, (body.nome_negocio, body.email, senha_hash, body.plano))
    return cliente

@app.patch("/admin/clientes/{cliente_id}/plano")
def admin_trocar_plano(cliente_id: int, body: AdminPlanoBody, conn=Depends(get_db), _admin=Depends(check_admin)):
    db_exec(conn, "UPDATE clientes SET plano = %s WHERE id = %s", (body.plano, cliente_id))
    return {"ok": True}

@app.patch("/admin/clientes/{cliente_id}/ativo")
def admin_toggle_ativo(cliente_id: int, body: AdminAtivoBody, conn=Depends(get_db), _admin=Depends(check_admin)):
    db_exec(conn, "UPDATE clientes SET ativo = %s WHERE id = %s", (body.ativo, cliente_id))
    return {"ok": True}

@app.delete("/admin/clientes/{cliente_id}")
def admin_deletar_cliente(cliente_id: int, conn=Depends(get_db), _admin=Depends(check_admin)):
    db_exec(conn, "DELETE FROM clientes WHERE id = %s", (cliente_id,))
    return {"ok": True}

@app.post("/admin/clientes/{cliente_id}/numeros")
def admin_add_numero(cliente_id: int, body: AdminNumeroBody, conn=Depends(get_db), _admin=Depends(check_admin)):
    numero = normalizar_numero(body.numero)
    if db_one(conn, "SELECT id FROM numeros_autorizados WHERE numero = %s", (numero,)):
        raise HTTPException(status_code=400, detail="Número já cadastrado")
    reg = db_exec(conn, """
        INSERT INTO numeros_autorizados (cliente_id, numero, nome) VALUES (%s,%s,%s) RETURNING *
    """, (cliente_id, numero, body.nome))
    return reg

@app.delete("/admin/numeros/{numero_id}")
def admin_del_numero(numero_id: int, conn=Depends(get_db), _admin=Depends(check_admin)):
    db_exec(conn, "DELETE FROM numeros_autorizados WHERE id = %s", (numero_id,))
    return {"ok": True}

@app.get("/admin/conexao")
def admin_get_conexao(conn=Depends(get_db), _admin=Depends(check_admin)):
    conexao = db_one(conn, "SELECT * FROM conexao_bot WHERE id = 1")
    try:
        async def _status():
            async with httpx.AsyncClient(timeout=10) as client:
                r = await client.get(f"{BAILEYS_URL}/status")
                return r.json()
    except Exception:
        pass
    return conexao

@app.post("/admin/conexao/resetar")
async def admin_resetar_conexao(conn=Depends(get_db), _admin=Depends(check_admin)):
    async with httpx.AsyncClient(timeout=15) as client:
        try:
            await client.post(f"{BAILEYS_URL}/resetar")
        except Exception as e:
            raise HTTPException(status_code=502, detail=f"Falha ao resetar sessão no Baileys: {e}")
    db_exec(conn, "UPDATE conexao_bot SET status = 'trocando', atualizado_em = NOW() WHERE id = 1")
    return {"ok": True, "mensagem": "Sessão resetada. Acesse /whatsapp/qrcode para escanear o novo QR."}

@app.post("/admin/conexao/notificar-clientes")
async def admin_notificar_clientes(conn=Depends(get_db), _admin=Depends(check_admin)):
    conexao = db_one(conn, "SELECT * FROM conexao_bot WHERE id = 1")
    numeros = db_all(conn, "SELECT numero FROM numeros_autorizados WHERE ativo = TRUE")
    texto = (
        "📢 Aviso: este é o número oficial atualizado do sistema de estoque. "
        "Salve este contato para continuar registrando suas movimentações por aqui."
    )
    enviados = 0
    for n in numeros:
        await enviar_whatsapp(n["numero"], texto)
        enviados += 1
    return {"ok": True, "notificados": enviados}

@app.get("/admin/log-trocas")
def admin_log_trocas(conn=Depends(get_db), _admin=Depends(check_admin)):
    return db_all(conn, "SELECT * FROM log_troca_numero ORDER BY criado_em DESC")

# Chamado internamente pelo Baileys quando conecta com um número (ver webhook abaixo)
class ConexaoAtualizaBody(BaseModel):
    numero: str

@app.post("/admin/conexao/atualizar")
def conexao_atualizar(body: ConexaoAtualizaBody, conn=Depends(get_db)):
    numero = normalizar_numero(body.numero)
    atual = db_one(conn, "SELECT * FROM conexao_bot WHERE id = 1")
    if atual and atual.get("numero_atual") and atual["numero_atual"] != numero:
        db_exec(conn, "INSERT INTO log_troca_numero (numero_antigo, numero_novo) VALUES (%s,%s)",
                (atual["numero_atual"], numero))
    db_exec(conn, "UPDATE conexao_bot SET numero_atual = %s, status = 'conectado', atualizado_em = NOW() WHERE id = 1",
            (numero,))
    return {"ok": True}

# ═════════════════════════════════════════
#  PRODUTOS (painel do cliente)
# ═════════════════════════════════════════
@app.get("/produtos")
def listar_produtos(cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    return db_all(conn, "SELECT * FROM produtos WHERE cliente_id = %s AND ativo = TRUE ORDER BY nome",
                  (cliente["id"],))

@app.post("/produtos")
def criar_produto(body: ProdutoBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    return db_exec(conn, """
        INSERT INTO produtos (cliente_id, nome, sku, custo_unitario, preco_venda, estoque_atual, unidade)
        VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING *
    """, (cliente["id"], body.nome, body.sku, body.custo_unitario, body.preco_venda, body.estoque_atual, body.unidade))

@app.put("/produtos/{produto_id}")
def atualizar_produto(produto_id: int, body: ProdutoBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    db_exec(conn, """
        UPDATE produtos SET nome=%s, sku=%s, custo_unitario=%s, preco_venda=%s, estoque_atual=%s, unidade=%s
        WHERE id=%s AND cliente_id=%s
    """, (body.nome, body.sku, body.custo_unitario, body.preco_venda, body.estoque_atual, body.unidade,
          produto_id, cliente["id"]))
    return {"ok": True}

@app.delete("/produtos/{produto_id}")
def deletar_produto(produto_id: int, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    db_exec(conn, "UPDATE produtos SET ativo = FALSE WHERE id = %s AND cliente_id = %s", (produto_id, cliente["id"]))
    return {"ok": True}

# ═════════════════════════════════════════
#  MATÉRIAS-PRIMAS (painel do cliente)
# ═════════════════════════════════════════
@app.get("/materias-primas")
def listar_materias_primas(cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    return db_all(conn, "SELECT * FROM materias_primas WHERE cliente_id = %s AND ativo = TRUE ORDER BY nome",
                  (cliente["id"],))

@app.post("/materias-primas")
def criar_materia_prima(body: MateriaPrimaBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    return db_exec(conn, """
        INSERT INTO materias_primas (cliente_id, nome, sku, custo_unitario, estoque_atual, unidade, estoque_minimo)
        VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING *
    """, (cliente["id"], body.nome, body.sku, body.custo_unitario, body.estoque_atual, body.unidade, body.estoque_minimo))

@app.put("/materias-primas/{materia_id}")
def atualizar_materia_prima(materia_id: int, body: MateriaPrimaBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    db_exec(conn, """
        UPDATE materias_primas SET nome=%s, sku=%s, custo_unitario=%s, estoque_atual=%s, unidade=%s, estoque_minimo=%s
        WHERE id=%s AND cliente_id=%s
    """, (body.nome, body.sku, body.custo_unitario, body.estoque_atual, body.unidade, body.estoque_minimo,
          materia_id, cliente["id"]))
    return {"ok": True}

@app.delete("/materias-primas/{materia_id}")
def deletar_materia_prima(materia_id: int, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    db_exec(conn, "UPDATE materias_primas SET ativo = FALSE WHERE id = %s AND cliente_id = %s",
            (materia_id, cliente["id"]))
    return {"ok": True}

@app.post("/materias-primas/movimentacao")
def criar_movimentacao_materia_prima_manual(body: MovimentacaoMateriaPrimaBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    materia = db_one(conn, "SELECT * FROM materias_primas WHERE id = %s AND cliente_id = %s",
                      (body.materia_prima_id, cliente["id"]))
    if not materia:
        raise HTTPException(status_code=404, detail="Matéria-prima não encontrada")
    _, alerta = aplicar_movimentacao_materia_prima(conn, cliente["id"], body.materia_prima_id, None, body.tipo,
                                                    body.quantidade, body.valor_unitario, origem="manual_painel")
    return {"ok": True, "alerta": alerta}

@app.get("/movimentacoes-materia-prima")
def listar_movimentacoes_materia_prima(
    materia_prima_id: Optional[int] = None, tipo: Optional[str] = None,
    de: Optional[str] = None, ate: Optional[str] = None,
    cliente=Depends(get_current_cliente), conn=Depends(get_db)
):
    sql = """
        SELECT mv.*, m.nome AS materia_prima_nome, n.nome AS registrado_por, p.nome AS produto_origem_nome
        FROM movimentacoes_materia_prima mv
        JOIN materias_primas m ON m.id = mv.materia_prima_id
        LEFT JOIN numeros_autorizados n ON n.id = mv.numero_autorizado_id
        LEFT JOIN produtos p ON p.id = mv.produto_id_origem
        WHERE mv.cliente_id = %s
    """
    params = [cliente["id"]]
    if materia_prima_id:
        sql += " AND mv.materia_prima_id = %s"; params.append(materia_prima_id)
    if tipo:
        sql += " AND mv.tipo = %s"; params.append(tipo)
    if de:
        sql += " AND mv.criado_em::date >= %s"; params.append(de)
    if ate:
        sql += " AND mv.criado_em::date <= %s"; params.append(ate)
    sql += " ORDER BY mv.criado_em DESC LIMIT 500"
    return db_all(conn, sql, tuple(params))

# ═════════════════════════════════════════
#  RECEITA / FICHA TÉCNICA (produto ⇄ matérias-primas)
# ═════════════════════════════════════════
@app.get("/produtos/{produto_id}/receita")
def obter_receita(produto_id: int, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s AND cliente_id = %s", (produto_id, cliente["id"]))
    if not produto:
        raise HTTPException(status_code=404, detail="Produto não encontrado")
    return db_all(conn, """
        SELECT r.id, r.materia_prima_id, r.quantidade_necessaria,
               m.nome AS materia_prima_nome, m.unidade, m.estoque_atual, m.custo_unitario
        FROM receita_itens r
        JOIN materias_primas m ON m.id = r.materia_prima_id
        WHERE r.produto_id = %s
        ORDER BY m.nome
    """, (produto_id,))

@app.put("/produtos/{produto_id}/receita")
def salvar_receita(produto_id: int, body: ReceitaBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s AND cliente_id = %s", (produto_id, cliente["id"]))
    if not produto:
        raise HTTPException(status_code=404, detail="Produto não encontrado")
    db_exec(conn, "DELETE FROM receita_itens WHERE produto_id = %s", (produto_id,))
    for item in body.itens:
        materia = db_one(conn, "SELECT id FROM materias_primas WHERE id = %s AND cliente_id = %s",
                          (item.materia_prima_id, cliente["id"]))
        if not materia:
            raise HTTPException(status_code=400, detail=f"Matéria-prima {item.materia_prima_id} não encontrada")
        db_exec(conn, """
            INSERT INTO receita_itens (produto_id, materia_prima_id, quantidade_necessaria)
            VALUES (%s,%s,%s)
        """, (produto_id, item.materia_prima_id, item.quantidade_necessaria))
    return {"ok": True}

# ═════════════════════════════════════════
#  MOVIMENTAÇÕES / DASHBOARD
# ═════════════════════════════════════════
@app.get("/movimentacoes")
def listar_movimentacoes(
    produto_id: Optional[int] = None, tipo: Optional[str] = None,
    de: Optional[str] = None, ate: Optional[str] = None,
    cliente=Depends(get_current_cliente), conn=Depends(get_db)
):
    sql = """
        SELECT m.*, p.nome AS produto_nome, n.nome AS registrado_por
        FROM movimentacoes m
        JOIN produtos p ON p.id = m.produto_id
        LEFT JOIN numeros_autorizados n ON n.id = m.numero_autorizado_id
        WHERE m.cliente_id = %s
    """
    params = [cliente["id"]]
    if produto_id:
        sql += " AND m.produto_id = %s"; params.append(produto_id)
    if tipo:
        sql += " AND m.tipo = %s"; params.append(tipo)
    if de:
        sql += " AND m.criado_em::date >= %s"; params.append(de)
    if ate:
        sql += " AND m.criado_em::date <= %s"; params.append(ate)
    sql += " ORDER BY m.criado_em DESC LIMIT 500"
    return db_all(conn, sql, tuple(params))

@app.post("/movimentacoes")
def criar_movimentacao_manual(body: MovimentacaoManualBody, cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    produto = db_one(conn, "SELECT * FROM produtos WHERE id = %s AND cliente_id = %s", (body.produto_id, cliente["id"]))
    if not produto:
        raise HTTPException(status_code=404, detail="Produto não encontrado")
    _, alertas = aplicar_movimentacao(conn, cliente["id"], body.produto_id, None, body.tipo, body.quantidade,
                                       body.valor_unitario, origem="manual_admin")
    return {"ok": True, "alertas": alertas}

@app.get("/dashboard/resumo")
def dashboard_resumo(cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    produtos = db_all(conn, "SELECT * FROM produtos WHERE cliente_id = %s AND ativo = TRUE", (cliente["id"],))
    valor_em_estoque = sum(float(p["estoque_atual"]) * float(p["custo_unitario"]) for p in produtos)

    hoje = datetime.utcnow().date()
    inicio_mes = hoje.replace(day=1)
    vendas_mes = db_one(conn, """
        SELECT COALESCE(SUM(valor_total),0) total, COALESCE(SUM(quantidade),0) qtd
        FROM movimentacoes WHERE cliente_id = %s AND tipo = 'venda' AND criado_em::date >= %s
    """, (cliente["id"], inicio_mes))
    custo_vendido_mes = db_all(conn, """
        SELECT m.produto_id, m.quantidade, p.custo_unitario
        FROM movimentacoes m JOIN produtos p ON p.id = m.produto_id
        WHERE m.cliente_id = %s AND m.tipo = 'venda' AND m.criado_em::date >= %s
    """, (cliente["id"], inicio_mes))
    custo_total_mes = sum(float(r["quantidade"]) * float(r["custo_unitario"]) for r in custo_vendido_mes)
    margem_mes = float(vendas_mes["total"]) - custo_total_mes

    return {
        "total_produtos": len(produtos),
        "valor_em_estoque": round(valor_em_estoque, 2),
        "produtos_estoque_baixo": [p for p in produtos if float(p["estoque_atual"]) <= 0],
        "vendas_mes_total": round(float(vendas_mes["total"]), 2),
        "vendas_mes_qtd": float(vendas_mes["qtd"]),
        "margem_mes": round(margem_mes, 2),
    }

@app.get("/dashboard/export.csv")
def export_csv(de: Optional[str] = None, ate: Optional[str] = None,
                cliente=Depends(get_current_cliente), conn=Depends(get_db)):
    sql = """
        SELECT m.criado_em, p.nome AS produto, m.tipo, m.quantidade, m.valor_unitario, m.valor_total, m.origem
        FROM movimentacoes m JOIN produtos p ON p.id = m.produto_id
        WHERE m.cliente_id = %s
    """
    params = [cliente["id"]]
    if de:
        sql += " AND m.criado_em::date >= %s"; params.append(de)
    if ate:
        sql += " AND m.criado_em::date <= %s"; params.append(ate)
    sql += " ORDER BY m.criado_em DESC"
    linhas = db_all(conn, sql, tuple(params))

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["Data", "Produto", "Tipo", "Quantidade", "Valor Unitário", "Valor Total", "Origem"])
    for l in linhas:
        writer.writerow([l["criado_em"], l["produto"], l["tipo"], l["quantidade"],
                          l["valor_unitario"], l["valor_total"], l["origem"]])
    buf.seek(0)
    return StreamingResponse(
        iter([buf.getvalue()]), media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=movimentacoes.csv"}
    )

# ═════════════════════════════════════════
#  WHATSAPP — webhook e status
# ═════════════════════════════════════════
@app.post("/webhook/mensagem")
async def webhook_mensagem(payload: dict, conn=Depends(get_db)):
    remote_jid = payload.get("remoteJid", "")
    numero_resolvido = payload.get("numeroResolvido")  # ex: "5511999998888@s.whatsapp.net", quando o remote_jid é @lid
    texto = payload.get("text")
    from_me = payload.get("fromMe", False)
    if from_me or not texto or remote_jid.endswith("@g.us"):
        return {"ok": True}

    # Autorização é checada pelo número de telefone real (resolvido do LID
    # quando disponível). A resposta, porém, SEMPRE vai pro remote_jid
    # original — trocar de endereço no meio da conversa quebra a sessão de
    # criptografia do WhatsApp e a mensagem fica "Aguardando mensagem".
    numero = normalizar_numero(numero_resolvido or remote_jid)
    numero_autorizado = buscar_numero_autorizado(conn, numero)
    if not numero_autorizado:
        await enviar_whatsapp(remote_jid, "❌ Número não autorizado. Fale com o administrador do sistema.")
        return {"ok": True}

    cliente = db_one(conn, "SELECT * FROM clientes WHERE id = %s AND ativo = TRUE", (numero_autorizado["cliente_id"],))
    if not cliente:
        await enviar_whatsapp(remote_jid, "❌ Conta inativa. Fale com o administrador.")
        return {"ok": True}

    resposta = await processar_texto(conn, cliente, numero_autorizado, texto)
    await enviar_whatsapp(remote_jid, resposta)
    return {"ok": True}

@app.get("/whatsapp/status")
async def whatsapp_status():
    async with httpx.AsyncClient(timeout=10) as client:
        try:
            r = await client.get(f"{BAILEYS_URL}/status")
            return r.json()
        except Exception:
            return {"connected": False}

@app.get("/whatsapp/qrcode")
async def whatsapp_qrcode():
    async with httpx.AsyncClient(timeout=10) as client:
        try:
            r = await client.get(f"{BAILEYS_URL}/qrcode-raw")
            return r.json()
        except Exception:
            return {"qr": None}

@app.get("/")
def root():
    return FileResponse("frontend/login.html")

@app.get("/login.html")
def pagina_login():
    return FileResponse("frontend/login.html")

@app.get("/admin-login.html")
def pagina_admin_login():
    return FileResponse("frontend/admin-login.html")

@app.get("/dashboard.html")
def pagina_dashboard():
    return FileResponse("frontend/dashboard.html")

@app.get("/admin.html")
def pagina_admin():
    return FileResponse("frontend/admin.html")
