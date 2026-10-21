"""Validação e formatação dos campos do formulário (lado servidor)."""
import re

REGIOES_PERMITIDAS = {"4ª Região"}

_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def so_digitos(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def limpar_texto(valor: str, multilinha: bool = False) -> str:
    """Remove caracteres de controle (quebram o XML do .docx) e normaliza espaços."""
    valor = _CONTROL.sub("", valor or "").replace("\r\n", "\n").replace("\r", "\n")
    if multilinha:
        linhas = [re.sub(r"[ \t]+", " ", ln).strip() for ln in valor.split("\n")]
        valor = "\n".join(linhas)
        return re.sub(r"\n{3,}", "\n\n", valor).strip()
    return re.sub(r"\s+", " ", valor).strip()


def cpf_valido(digitos: str) -> bool:
    if len(digitos) != 11 or digitos == digitos[0] * 11:
        return False
    for n in (9, 10):
        soma = sum(int(digitos[i]) * (n + 1 - i) for i in range(n))
        if (soma * 10 % 11) % 10 != int(digitos[n]):
            return False
    return True


def formatar_cpf(d: str) -> str:
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"


def formatar_cep(d: str) -> str:
    return f"{d[:5]}-{d[5:]}"


def formatar_telefone(d: str) -> str:
    if len(d) == 11:
        return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    return f"({d[:2]}) {d[2:6]}-{d[6:]}"


def validar_formulario(form) -> tuple[dict, dict]:
    """Retorna (dados_limpos, erros)."""
    erros: dict[str, str] = {}
    dados: dict[str, str] = {}

    regiao = limpar_texto(form.get("regiao", ""))
    if regiao not in REGIOES_PERMITIDAS:
        erros["regiao"] = "Selecione a região do Conselho Popular."
    dados["regiao"] = regiao

    nome = limpar_texto(form.get("nome", ""))
    if len(nome) < 5 or " " not in nome:
        erros["nome"] = "Informe o nome completo."
    elif len(nome) > 120:
        erros["nome"] = "Nome muito longo (máx. 120 caracteres)."
    dados["nome"] = nome

    cpf = so_digitos(form.get("cpf", ""))
    if not cpf_valido(cpf):
        erros["cpf"] = "CPF inválido."
    dados["cpf"] = formatar_cpf(cpf) if len(cpf) == 11 else cpf

    cep = so_digitos(form.get("cep", ""))
    if len(cep) != 8:
        erros["cep"] = "CEP inválido."
    dados["cep"] = formatar_cep(cep) if len(cep) == 8 else cep

    endereco = limpar_texto(form.get("endereco", "")).strip(" ,")
    if len(endereco) < 3:
        erros["endereco"] = "Informe o endereço completo."
    elif len(endereco) > 200:
        erros["endereco"] = "Endereço muito longo (máx. 200 caracteres)."
    dados["endereco"] = endereco

    bairro = limpar_texto(form.get("bairro", ""))
    if len(bairro) < 2:
        erros["bairro"] = "Informe o bairro."
    elif len(bairro) > 100:
        erros["bairro"] = "Bairro muito longo (máx. 100 caracteres)."
    dados["bairro"] = bairro

    tel = so_digitos(form.get("telefone", ""))
    if len(tel) not in (10, 11) or (len(tel) == 11 and tel[2] != "9"):
        erros["telefone"] = "Telefone inválido. Use DDD + número."
    dados["telefone"] = formatar_telefone(tel) if len(tel) in (10, 11) else tel

    demanda = limpar_texto(form.get("demanda", ""), multilinha=True)
    if len(demanda) < 10:
        erros["demanda"] = "Descreva sua demanda (mín. 10 caracteres)."
    elif len(demanda) > 5000:
        erros["demanda"] = "Texto muito longo (máx. 5000 caracteres)."
    dados["demanda"] = demanda

    protocolo = limpar_texto(form.get("protocolo", ""))
    if len(protocolo) > 60:
        erros["protocolo"] = "Protocolo muito longo (máx. 60 caracteres)."
    dados["protocolo"] = protocolo

    return dados, erros
