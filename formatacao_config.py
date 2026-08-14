from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENTRADA_DIR = BASE_DIR / "divididos"
SAIDA_DIR = BASE_DIR / "saida"

DEBUG_DIR = ENTRADA_DIR / "debug_ia"

GEMINI_FLASH_LITE_MODEL = "gemini-3.5-flash-lite"
GEMINI_FLASH_MODEL = "gemini-3.6-flash"
GEMINI_3_FLASH_MODEL = "gemini-3.1-flash-lite"
GEMINI_TEMPERATURE = 0

# Regras de formatação de traduções
ROLE_STYLES = {
    "registered_name": {"bold": True, "uppercase": True},
    "spouse_name": {"bold": True, "uppercase": True},
    "maiden_name": {"bold": True, "uppercase": True},
    "married_name": {"bold": True, "uppercase": True},

    "father_name": {"bold": False, "uppercase": True},
    "mother_name": {"bold": False, "uppercase": True},
    "grandparent_name": {"bold": False, "uppercase": True},
    "father_in_law": {"bold": False, "uppercase": True},
    "mother_in_law": {"bold": False, "uppercase": True},
    "grandparent_spouse_name": {"bold": False, "uppercase": True},

    "birth_date": {"bold": True, "uppercase": False},
    "marriage_date": {"bold": True, "uppercase": False},
    "death_date": {"bold": True, "uppercase": False},

    "birth_place": {"bold": True, "uppercase": False},
    "marriage_place": {"bold": True, "uppercase": False},
    "death_place": {"bold": True, "uppercase": False},

    "annotation_title": {"bold": True, "uppercase": False},
    "annotation_subject": {"bold": True, "uppercase": False},
    "rectification_title": {"bold": True, "uppercase": False},
    "rectification_old_name": {"bold": False, "uppercase": True},
    "rectification_new_name": {"bold": True, "uppercase": True},

    "witness": {"bold": False, "uppercase": False, "capitalize_each_word": True},
    "officials": {"bold": False, "uppercase": False, "capitalize_each_word": True},
    "employees": {"bold": False, "uppercase": False, "capitalize_each_word": True},
}

ROLE_DEFINITIONS_GERAIS = {
    "registered_name": "Nome da pessoa registrada na certidão.",
    "spouse_name": "Nome do cônjuge na certidão de casamento.",
    "maiden_name": "Nome de solteiro(a) / nome anterior antes do casamento.",
    "married_name": "Nome após casamento.",
    "father_name": "Nome do pai.",
    "mother_name": "Nome da mãe.",
    "grandparent_name": "Nome de avô/avó.",
    "birth_date": "Data de nascimento.",
    "marriage_date": "Data de casamento.",
    "registration_date": "Data de registro / lavratura / transcrição.",
    "death_date": "Data de óbito.",
    "birth_place": "Cidade de nascimento.",
    "marriage_place": "Cidade do casamento.",
    "death_place": "Cidade do óbito.",
    "annotation_title": "Título de anotação/averbação/observação.",
    "annotation_subject": "Trecho descritivo após um título de anotação quando o título aparece isolado.",
    "rectification_title": "Título de retificação.",
    "rectification_old_name": "Nome antigo na retificação.",
    "rectification_new_name": "Nome novo na retificação.",
    "witness": "Testemunha do ato / registro",
    "officials": "Oficial que assinou a certidão / registro",
    "employees": "Funcionário que assinou a certidão / registro",
}

ROLE_DEFINITIONS_NASCIMENTO = {
    "declarant": "Nome do declarante do nascimento.",
    "registration_place": "Local onde foi registrado / lavrado / transcrito."
}

ROLE_DEFINITIONS_CASAMENTO = {
    "father_in_law": "Nome do pai do cônjuge.",
    "mother_in_law": "Nome da mãe do cônjuge.",
    "grandparent_spouse_name": "Nome de avô/avó do cônjuge."
}

ROLE_DEFINITIONS_POR_TIPO = {
    "nascimento": {
        **ROLE_DEFINITIONS_GERAIS,
        **ROLE_DEFINITIONS_NASCIMENTO,
    },
    "casamento": {
        **ROLE_DEFINITIONS_GERAIS,
        **ROLE_DEFINITIONS_CASAMENTO,
    },
    "obito": ROLE_DEFINITIONS_GERAIS,
    "desconhecido": ROLE_DEFINITIONS_GERAIS,
}

ROLE_ALTERACAO_PRIMEIRA_OCORRENCIA = {
    "birth_date",
    "marriage_date",
    "death_date",
    "birth_place",
    "marriage_place",
    "death_place",
}

SYSTEM_PROMPT = """
Você é um identificador de trechos.
Você receberá uma lista de parágrafos de traduções juramentadas de certidões em italiano.
Cada parágrafo possui um id.
Sua única tarefa é identificar quais trechos precisam de formatação.
Não traduza.
Não altere nenhum texto.
Não invente texto.
Cada trecho informado deve existir exatamente dentro do parágrafo.
Responda APENAS com JSON.
Formato:
{{
    "segments":[
        {{
            "id":1,
            "text":"Augustinho Barbieri",
            "role":"registered_name"
        }}
    ]
}}

Retorne cada trecho literal apenas uma vez por parágrafo.
Se o mesmo texto aparecer várias vezes no parágrafo, retorne apenas uma ocorrência.
Não converta datas por extenso em datas numéricas.
Não reescreva nomes.
Não atribua mais de uma role ao mesmo texto literal.
O texto deve ser copiado exatamente como aparece no parágrafo.

Roles disponíveis:
{roles}
""".strip()