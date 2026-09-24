from core.database import DatabaseManager
from core.email_sender import EmailSender

db = DatabaseManager('data/escala.db')
sender = EmailSender(test_mode=True)
df_last = db.get_last_grade_df()

with db.get_connection() as conn:
    c = conn.cursor()
    c.execute("SELECT funcionario, data_escala, tipo, campo, valor_antigo, valor_novo FROM changes WHERE tipo = 'ALTERACAO' LIMIT 10")
    rows = c.fetchall()

if rows:
    func_name = rows[0][0]
    func_changes = [
        {
            'grade_id': 1,
            'funcionario': r[0],
            'data': r[1],
            'tipo': r[2],
            'campo': r[3],
            'valor_antigo': r[4],
            'valor_novo': r[5]
        } for r in rows if r[0] == func_name
    ]
    df_func = df_last[df_last['funcionario'].str.lower() == func_name.lower()] if not df_last.empty and 'funcionario' in df_last.columns else None
    html = sender._generate_html_body(func_name, func_changes, df_func)
    
    with open('preview_mudanca_horario.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Preview de mudança de horário gerado com sucesso para {func_name}: preview_mudanca_horario.html')
