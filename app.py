from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import sqlite3
import os
import io
from datetime import datetime, timedelta
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from apscheduler.schedulers.background import BackgroundScheduler
import sys 


base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))


template_folder_path = os.path.join(base_path, 'templates')
database_folder_path = os.path.join(base_path, 'database')
database_file_path = os.path.join(database_folder_path, 'protocolos.db')


app = Flask(__name__, template_folder=template_folder_path)


os.makedirs(database_folder_path, exist_ok=True)


DATABASE = database_file_path


EMAIL_ORIGEM = 'cgmprotocolo@gmail.com'
SENHA_EMAIL = 'zvjt ohkr zzci'
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS processos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_processo TEXT NOT NULL,
            volume TEXT,
            secretaria TEXT NOT NULL,
            data_entrada TEXT NOT NULL,
            hora_entrada TEXT NOT NULL,
            data_saida TEXT,
            hora_saida TEXT,
            destino TEXT,
            genero TEXT NOT NULL,
            especie TEXT NOT NULL,
            objeto TEXT NOT NULL,
            contratada TEXT,
            recorrente TEXT DEFAULT 'NÃO',
            prioridade TEXT NOT NULL,
            tecnico TEXT,
            data_analise TEXT,
            numero_despacho TEXT,
            observacao TEXT,
            aviso_enviado INTEGER DEFAULT 0
        );
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS process_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            process_id INTEGER NOT NULL,
            field_name TEXT NOT NULL,
            old_value TEXT,
            new_value TEXT,
            changed_at TEXT NOT NULL,
            changed_by TEXT DEFAULT 'Sistema',
            FOREIGN KEY (process_id) REFERENCES processos(id) ON DELETE CASCADE
        );
    ''')
    conn.commit()
    conn.close()


def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def enviar_email(destinatario, numero_processo, prazo):
    assunto = f"[Alerta] Prazo do processo {numero_processo}"
    corpo = f"""
    Alerta de prazo:
    
    Processo: {numero_processo}
    Data limite: {prazo}
    Status: {'ÚLTIMO DIA!' if (datetime.strptime(prazo, "%Y-%m-%d").date() - datetime.now().date()).days == 0 else 'Falta 1 dia'}
    
    Acesse o sistema para mais detalhes.
    """
    
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ORIGEM
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))
    
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ORIGEM, SENHA_EMAIL)
            server.send_message(msg)
            print(f"Email enviado para {destinatario} sobre o processo {numero_processo}")
            return True
    except Exception as e:
        print(f"Erro ao enviar email: {e}")
        return False
    

def verificar_prazos():
    print("Verificando prazos...")
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT id, numero_processo, data_entrada, prioridade,
            CASE prioridade
                WHEN 'URGENTE' THEN date(data_entrada, '+1 day')
                WHEN 'ALTA' THEN date(data_entrada, '+2 day')
                WHEN 'MÉDIA' THEN date(data_entrada, '+5 day')
                WHEN 'BAIXA' THEN date(data_entrada, '+7 day')
                ELSE date(data_entrada, '+999 day')
            END as data_prazo
        FROM processos
        WHERE (data_saida IS NULL OR data_saida = '')
        AND (aviso_enviado = 0)
    ''')

    processos = cursor.fetchall()
    hoje = datetime.now().date()

    for processo in processos:
        data_prazo = datetime.strptime(processo["data_prazo"], "%Y-%m-%d").date()
        dias_restantes = (data_prazo - hoje).days

        if dias_restantes <= 1:
            enviar_email(EMAIL_ORIGEM, processo["numero_processo"], processo["data_prazo"])
            cursor.execute("UPDATE processos SET aviso_enviado = 1 WHERE id = ?", (processo["id"],))
            conn.commit()

    conn.close()

def calcular_prazo(data_entrada, prioridade):
    if not data_entrada or not prioridade:
        return "-"
    try:
        data = datetime.strptime(data_entrada, '%Y-%m-%d')
        dias = {'BAIXA': 7, 'MÉDIA': 5, 'ALTA': 2, 'URGENTE': 1}.get(prioridade, 0)
        data_prazo = data + timedelta(days=dias)
        return data_prazo.strftime('%d/%m/%Y')
    except:
        return "-"



@app.route('/')
def index():
    return render_template('formulario.html')

@app.route('/processo/<int:process_id>/historico')
def ver_historico_processo(process_id):
    conn = get_db()
    cursor = conn.cursor()
    
    
    cursor.execute('SELECT numero_processo FROM processos WHERE id = ?', (process_id,))
    process_info = cursor.fetchone()
    process_number = process_info['numero_processo'] if process_info else 'N/A'

   
    cursor.execute('''
        SELECT field_name, old_value, new_value, changed_at, changed_by
        FROM process_history
        WHERE process_id = ?
        ORDER BY changed_at DESC
    ''', (process_id,))
    
    history_records = cursor.fetchall()
    conn.close()
    
   
    formatted_history = []
    for record in history_records:
        rec_dict = dict(record)
        if rec_dict['changed_at']:
            
            rec_dict['changed_at_formatted'] = datetime.strptime(rec_dict['changed_at'], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y %H:%M:%S')
        formatted_history.append(rec_dict)

    return render_template('historico_processo.html', 
                           history=formatted_history, 
                           process_id=process_id, 
                           process_number=process_number)


@app.route('/salvar', methods=['POST'])
def salvar():
    dados = request.get_json()
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute('''
        INSERT INTO processos (
            numero_processo, volume, secretaria, data_entrada, hora_entrada,
            data_saida, hora_saida, destino, genero, especie, objeto,
            contratada, recorrente, prioridade, tecnico, data_analise,
            numero_despacho, observacao
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dados['numero_processo'], dados.get('volume'),
            dados['secretaria'], dados['data_entrada'], dados['hora_entrada'],
            dados.get('data_saida'), dados.get('hora_saida'),
            dados.get('destino'), dados['genero'], dados['especie'],
            dados['objeto'], dados.get('contratada'), dados.get('recorrente'),
            dados['prioridade'], dados.get('tecnico'), dados.get('data_analise'),
            dados.get('numero_despacho'), dados.get('observacao')
        ))
        conn.commit()
        return jsonify({"success": True, "message": "Processo salvo!"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})
    finally:
        conn.close()

@app.route('/listar')
def listar_processos():
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
    SELECT *, 
        CASE prioridade
            WHEN 'URGENTE' THEN date(data_entrada, '+1 day')
            WHEN 'ALTA' THEN date(data_entrada, '+2 days')
            WHEN 'MÉDIA' THEN date(data_entrada, '+5 days')
            WHEN 'BAIXA' THEN date(data_entrada, '+7 days')
            ELSE date(data_entrada, '+999 days')
        END as data_prazo_calculado_backend
    FROM processos 
    WHERE data_saida IS NULL OR data_saida = '' 
    ORDER BY data_prazo_calculado_backend ASC, prioridade DESC
    ''')
    
    processos = cursor.fetchall()
    conn.close()
    
    processos_dict = []
    for processo in processos:
        p_dict = dict(processo)
        if p_dict['data_prazo_calculado_backend']:
            try:
                data_prazo_obj = datetime.strptime(p_dict['data_prazo_calculado_backend'], "%Y-%m-%d").date()
                hoje = datetime.now().date()
                dias_restantes = (data_prazo_obj - hoje).days
                p_dict['prazo_formatado'] = f"{dias_restantes} dia(s) {'atrasado' if dias_restantes < 0 else 'restante(s)'}"
            except ValueError:
                p_dict['prazo_formatado'] = "-"
        else:
            p_dict['prazo_formatado'] = "-"
        processos_dict.append(p_dict)
        
    return render_template('lista_processos.html', processos=processos_dict, prioridade_filtro='todas', termo_pesquisa='')

@app.route('/pesquisar', methods=['GET'])
def pesquisar_processos():
    termo = request.args.get('termo', '').strip()
    prioridade_filtro = request.args.get('prioridade', 'todas') 
    
    conn = get_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT *, 
            CASE prioridade
                WHEN 'URGENTE' THEN date(data_entrada, '+1 day')
                WHEN 'ALTA' THEN date(data_entrada, '+2 days')
                WHEN 'MÉDIA' THEN date(data_entrada, '+5 days')
                WHEN 'BAIXA' THEN date(data_entrada, '+7 days')
                ELSE date(data_entrada, '+999 days')
            END as data_prazo_calculado_backend
        FROM processos 
        WHERE (data_saida IS NULL OR data_saida = '')
        AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)
    '''
    params = [f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%']

    if prioridade_filtro != 'todas':
        query += ' AND prioridade = ?'
        params.append(prioridade_filtro)
    
    query += ' ORDER BY data_prazo_calculado_backend ASC, prioridade DESC'
    
    cursor.execute(query, tuple(params))
    
    processos = cursor.fetchall()
    conn.close()
    
    processos_dict = []
    for processo in processos:
        p_dict = dict(processo)
        if p_dict['data_prazo_calculado_backend']:
            try:
                data_prazo_obj = datetime.strptime(p_dict['data_prazo_calculado_backend'], "%Y-%m-%d").date()
                hoje = datetime.now().date()
                dias_restantes = (data_prazo_obj - hoje).days
                p_dict['prazo_formatado'] = f"{dias_restantes} dia(s) {'atrasado' if dias_restantes < 0 else 'restante(s)'}"
            except ValueError:
                p_dict['prazo_formatado'] = "-"
        else:
            p_dict['prazo_formatado'] = "-"
        processos_dict.append(p_dict)
        
    return render_template('lista_processos.html', processos=processos_dict, prioridade_filtro=prioridade_filtro, termo_pesquisa=termo)

@app.route('/listar_por_prioridade/<prioridade>')
def listar_por_prioridade(prioridade):
    termo = request.args.get('termo', '').strip()
    conn = get_db()
    cursor = conn.cursor()
    
    if prioridade.lower() == 'todas':
        query = '''
            SELECT *, 
                CASE prioridade
                    WHEN 'URGENTE' THEN date(data_entrada, '+1 day')
                    WHEN 'ALTA' THEN date(data_entrada, '+2 days')
                    WHEN 'MÉDIA' THEN date(data_entrada, '+5 days')
                    WHEN 'BAIXA' THEN date(data_entrada, '+7 days')
                    ELSE date(data_entrada, '+999 days')
                END as data_prazo_calculado_backend
            FROM processos 
            WHERE (data_saida IS NULL OR data_saida = '')
        '''
        params = []
    else:
        query = '''
            SELECT *, 
                CASE prioridade
                    WHEN 'URGENTE' THEN date(data_entrada, '+1 day')
                    WHEN 'ALTA' THEN date(data_entrada, '+2 days')
                    WHEN 'MÉDIA' THEN date(data_entrada, '+5 days')
                    WHEN 'BAIXA' THEN date(data_entrada, '+7 days')
                    ELSE date(data_entrada, '+999 days')
                END as data_prazo_calculado_backend
            FROM processos 
            WHERE prioridade = ? AND (data_saida IS NULL OR data_saida = '')
        '''
        params = [prioridade]

    if termo:
        query += ' AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)'
        params.extend([f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%'])
    
    query += ' ORDER BY data_prazo_calculado_backend ASC, prioridade DESC'
    
    cursor.execute(query, tuple(params))
    
    processos = cursor.fetchall()
    conn.close()
    
    processos_dict = []
    for processo in processos:
        p_dict = dict(processo)
        if p_dict['data_prazo_calculado_backend']:
            try:
                data_prazo_obj = datetime.strptime(p_dict['data_prazo_calculado_backend'], "%Y-%m-%d").date()
                hoje = datetime.now().date()
                dias_restantes = (data_prazo_obj - hoje).days
                p_dict['prazo_formatado'] = f"{dias_restantes} dia(s) {'atrasado' if dias_restantes < 0 else 'restante(s)'}"
            except ValueError:
                p_dict['prazo_formatado'] = "-"
        else:
            p_dict['prazo_formatado'] = "-"
        processos_dict.append(p_dict)
        
    return render_template('lista_processos.html', processos=processos_dict, prioridade_filtro=prioridade, termo_pesquisa=termo)

@app.route('/editar/<int:id>', methods=['GET'])
def editar_processo(id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM processos WHERE id = ?', (id,))
    processo = cursor.fetchone()
    conn.close()
    
    if processo:
        processo_dict = dict(processo)
        for key in processo_dict:
            if processo_dict[key] is None:
                processo_dict[key] = ''
        return render_template('editar_processo.html', processo=processo_dict)
    return "Processo não encontrado", 404

@app.route('/finalizados')
def listar_finalizados():
    conn = get_db()
    cursor = conn.cursor()
    
    data_saida_filtro = request.args.get('data_saida', '').strip()
    prioridade_filtro = request.args.get('prioridade', 'todas')
    termo = request.args.get('termo', '').strip()

    query = "SELECT * FROM processos WHERE data_saida IS NOT NULL AND data_saida != ''"
    params = []

    if prioridade_filtro != 'todas':
        query += ' AND prioridade = ?'
        params.append(prioridade_filtro)

    if data_saida_filtro:
        query += ' AND data_saida = ?'
        params.append(data_saida_filtro)

    if termo:
        query += ' AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)'
        params.extend([f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%'])
    
    query += ' ORDER BY data_saida DESC'
    
    cursor.execute(query, tuple(params))
    processos = cursor.fetchall()
    conn.close()
    
    processos_dict = [dict(processo) for processo in processos]
    return render_template('finalizados.html', 
                           processos=processos_dict, 
                           prioridade_filtro=prioridade_filtro, 
                           termo_pesquisa=termo,
                           data_saida_filtro=data_saida_filtro)


@app.route('/finalizados_por_prioridade/<prioridade>')
def finalizados_por_prioridade(prioridade):
    termo = request.args.get('termo', '').strip()
    data_saida_filtro = request.args.get('data_saida', '').strip()
    conn = get_db()
    cursor = conn.cursor()
    
    query = "SELECT * FROM processos WHERE data_saida IS NOT NULL AND data_saida != ''"
    params = []

    if prioridade.lower() != 'todas':
        query += ' AND prioridade = ?'
        params.append(prioridade)

    if data_saida_filtro:
        query += ' AND data_saida = ?'
        params.append(data_saida_filtro)

    if termo:
        query += ' AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)'
        params.extend([f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%'])
    
    query += ' ORDER BY data_saida DESC'
    
    cursor.execute(query, tuple(params))
    
    processos = cursor.fetchall()
    conn.close()
    
    processos_dict = [dict(processo) for processo in processos]
    return render_template('finalizados.html', 
                           processos=processos_dict, 
                           prioridade_filtro=prioridade, 
                           termo_pesquisa=termo,
                           data_saida_filtro=data_saida_filtro)


@app.route('/pesquisar_finalizados', methods=['GET'])
def pesquisar_finalizados():
    termo = request.args.get('termo', '').strip()
    prioridade_filtro = request.args.get('prioridade', 'todas')
    data_saida_filtro = request.args.get('data_saida', '').strip()

    conn = get_db()
    cursor = conn.cursor()

    query = "SELECT * FROM processos WHERE data_saida IS NOT NULL AND data_saida != ''"
    params = []

    if prioridade_filtro != 'todas':
        query += ' AND prioridade = ?'
        params.append(prioridade_filtro)

    if data_saida_filtro:
        query += ' AND data_saida = ?'
        params.append(data_saida_filtro)

    if termo:
        query += ' AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)'
        params.extend([f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%'])

    query += ' ORDER BY data_saida DESC'

    cursor.execute(query, tuple(params))
    processos = cursor.fetchall()
    conn.close()

    processos_dict = [dict(processo) for processo in processos]
    return render_template('finalizados.html', 
                           processos=processos_dict, 
                           prioridade_filtro=prioridade_filtro, 
                           termo_pesquisa=termo,
                           data_saida_filtro=data_saida_filtro)


@app.route('/atualizar/<int:id>', methods=['POST'])
def atualizar_processo(id):
    dados = request.get_json()
    
    conn = get_db()
    cursor = conn.cursor()
    
    try:
        cursor.execute('SELECT * FROM processos WHERE id = ?', (id,))
        current_process_row = cursor.fetchone()
        
        if not current_process_row:
            return jsonify({"success": False, "message": "Processo não encontrado."}), 404
        
        current_process_data = dict(current_process_row)
        
        cursor.execute('''
        UPDATE processos SET
            numero_processo = ?,
            volume = ?,
            secretaria = ?,
            data_entrada = ?,
            hora_entrada = ?,
            data_saida = ?,
            hora_saida = ?,
            destino = ?,
            genero = ?,
            especie = ?,
            objeto = ?,
            contratada = ?,
            recorrente = ?,
            prioridade = ?,
            tecnico = ?,
            data_analise = ?,
            numero_despacho = ?,
            observacao = ?
        WHERE id = ?
        ''', (
            dados['numero_processo'], dados.get('volume'),
            dados['secretaria'], dados['data_entrada'], dados['hora_entrada'],
            dados.get('data_saida'), dados.get('hora_saida'),
            dados.get('destino'), dados['genero'], dados['especie'],
            dados['objeto'], dados.get('contratada'), dados.get('recorrente'),
            dados['prioridade'], dados.get('tecnico'), dados.get('data_analise'),
            dados.get('numero_despacho'), dados.get('observacao'),
            id
        ))
        
        changed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        changed_by = request.remote_addr
        
        fields_to_track = [
            'numero_processo', 'volume', 'secretaria', 'data_entrada', 'hora_entrada',
            'data_saida', 'hora_saida', 'destino', 'genero', 'especie', 'objeto',
            'contratada', 'recorrente', 'prioridade', 'tecnico', 'data_analise',
            'numero_despacho', 'observacao'
        ]

        for field in fields_to_track:
            old_val = str(current_process_data.get(field) or '')
            new_val = str(dados.get(field) or '')

            if old_val != new_val:
                cursor.execute('''
                    INSERT INTO process_history (process_id, field_name, old_value, new_value, changed_at, changed_by)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (id, field, old_val, new_val, changed_at, changed_by))
        
        conn.commit()
        return jsonify({"success": True, "message": "Processo atualizado!"})
    
    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "message": str(e)})
    
    finally:
        conn.close()
        

@app.route('/exportar_finalizados_excel', methods=['GET'])
def exportar_finalizados_excel():
    data_saida_filtro = request.args.get('data_saida', '').strip()
    prioridade_filtro = request.args.get('prioridade', 'todas')
    termo = request.args.get('termo', '').strip()

    if not data_saida_filtro:
        return jsonify({"success": False, "message": "Por favor, selecione uma Data de Saída para exportar os processos do dia."}), 400

    conn = get_db()
    cursor = conn.cursor()

    
    query = "SELECT numero_processo, volume, secretaria, data_entrada, hora_entrada, data_saida, hora_saida, destino, genero, especie, objeto, contratada, recorrente, prioridade, tecnico, data_analise, numero_despacho, observacao FROM processos WHERE data_saida IS NOT NULL AND data_saida != ''"
    params = []

    if prioridade_filtro != 'todas':
        query += ' AND prioridade = ?'
        params.append(prioridade_filtro)

    if data_saida_filtro:
        query += ' AND data_saida = ?'
        params.append(data_saida_filtro)

    if termo:
        query += ' AND (numero_processo LIKE ? OR secretaria LIKE ? OR objeto LIKE ? OR contratada LIKE ? OR tecnico LIKE ?)'
        params.extend([f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%', f'%{termo}%'])

    query += ' ORDER BY data_saida DESC'

    cursor.execute(query, tuple(params))
    processos = cursor.fetchall()
    conn.close()

    if not processos:
        return jsonify({"success": False, "message": "Nenhum processo finalizado encontrado para a data e filtros selecionados."}), 404

   
    df = pd.DataFrame(processos, columns=[col[0] for col in cursor.description])

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Processos Finalizados')
    writer.close() 
    output.seek(0) 


    filename = f"processos_finalizados_{data_saida_filtro}.xlsx" 

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        download_name=filename,
        as_attachment=True
    )


@app.route('/get_process_by_number/<string:numero_processo>', methods=['GET'])
def get_process_by_number(numero_processo):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM processos WHERE numero_processo = ?', (numero_processo,))
    processo = cursor.fetchone()
    conn.close()
    
    if processo:
        
        return jsonify(dict(processo))
    return jsonify({"message": "Processo não encontrado"}), 404


if __name__ == "__main__":
    init_db() 
    
    scheduler = BackgroundScheduler()
    scheduler.add_job(verificar_prazos, 'interval', minutes=5)
    scheduler.start()
    
    try:
        app.run(debug=True)
    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()