import sqlite3

def criar_banco():
    conn = sqlite3.connect('database/protocolos.db')
    cursor = conn.cursor()
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS processos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_processo TEXT UNIQUE,
        volume INTEGER,
        secretaria TEXT,
        data_entrada TEXT,
        hora_entrada TEXT,
        data_saida TEXT,
        hora_saida TEXT,
        destino TEXT,
        genero TEXT,
        especie TEXT,
        objeto TEXT,
        contratada TEXT,
        recorrente TEXT,
        prioridade TEXT,
        tecnico TEXT,
        data_analise TEXT,
        numero_despacho TEXT,
        observacao TEXT,
        aviso_enviado INTEGER DEFAULT 0
    )
    ''')
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    criar_banco()