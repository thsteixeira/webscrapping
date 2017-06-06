from bs4 import BeautifulSoup
import sqlite3
import os
import locale
from datetime import datetime

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR')
except:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil')

try:
    tipo_de_processo = "primeiro_grau"
    diretorio = "processos_descompressão_mai_2017"
    conn = sqlite3.connect(diretorio + '/processos_descompressão_mai_2017.db')
    cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS processos
                (id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero_processo TEXT UNIQUE, vara TEXT, dias_parados INTEGER, mensagem_download TEXT)''')
    cur.execute('''CREATE TABLE IF NOT EXISTS andamentos (numero_processo TEXT, data TEXT, 
                andamento TEXT, FOREIGN KEY(numero_processo) REFERENCES processos(numero_processo))''')

    for html in os.listdir(diretorio+"/html"):
        if html.endswith(".html") and html != "236812016.html":
            print(html)
            numero_processo = html.replace(".html", "")
            with open(diretorio+"/html/"+html) as leitura:
                page_soup = BeautifulSoup(leitura, 'html.parser')
            # numero_processo = page_soup.find(string="Número:").find_next("td")
            # print(numero_processo.text)
            if tipo_de_processo == "primeiro_grau":
                vara = page_soup.find(string="Vara:").find_next("td").text
            elif tipo_de_processo == "precatorio":
                vara = "Tribunal de Justiça"
            cur.execute("""INSERT INTO processos (numero_processo, vara, mensagem_download) VALUES
                        (?, ?, ?)""", (numero_processo, vara, "Consulta realizada com sucesso."))
            primeira_movimentacao = page_soup.find(class_="movimentacao")
            ultima_data = primeira_movimentacao.h1.text.replace("\n", "").replace("\t", "")
            qnt_dias = datetime.now() - datetime.strptime(ultima_data, "%A, %d de %B de %Y")
            qnt_dias = str(qnt_dias.days)
            cur.execute("""UPDATE processos SET
                        dias_parados = ? WHERE numero_processo = ?""", (qnt_dias, numero_processo))
            movimentacoes = page_soup.find_all(class_="movimentacao")
            for movimentacao in movimentacoes:
                data = movimentacao.h1.text.replace("\n", "").replace("\t", "")
                # print(movimentacao.h1.text.replace("\n", ""))
                movimentacao.h1.decompose()
                # movimentacao.h2.decompose()
                #movimentacao = movimentacao.text.replace("\n", "").replace("\t", "")
                movimentacao = ' '.join(movimentacao.text.split())
                #print(movimentacao)
                cur.execute("""INSERT INTO andamentos (numero_processo, data, andamento) VALUES 
                            (?, ?, ?)""", (numero_processo, data, movimentacao))
finally:
    conn.commit()
    conn.close()