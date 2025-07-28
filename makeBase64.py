import base64

# Замените путь на свой шаблон
template_path = "template_final.xlsx"

with open(template_path, "rb") as f:
    encoded = base64.b64encode(f.read()).decode("utf-8")

# Разбиваем на строки по 76 символов (как принято в base64-блоках)
lines = [encoded[i:i+76] for i in range(0, len(encoded), 76)]
formatted = '\n'.join(lines)

print('EXCEL_TEMPLATE_BASE64 = """')
print(formatted)
print('""".strip()')