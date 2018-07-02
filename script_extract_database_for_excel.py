import os, time, sys,zipfile,glob
import locale
import datetime 
import string as rp_directory
import mysql.connector  as msql
import psycopg2 as pg
import pandas as pd
import pandas.io.sql as psql

class extractDate ():																		        # Inicia a classe
	# reload(sys)																					# Recarrega a biblioteca do sistema
	# sys.setdefaultencoding("utf-8")																# insere uma nova codificacao de caracter unicode utf8
	try: 
		locale.setlocale(locale.LC_ALL, 'pt_BR') 
	except: 
		locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil') 
	
	def for_connect_postgres():																		# realiza conexao com o banco de dados
		try:
			connection = pg.connect(host='',database='',user='', password='')				# Dados de conexao com o banco
			return connection												    	        # retorna a conexao do banco.
		except Exception: 
			return ('Erro de Conexao: ',Exception) 									        # retorna um erro caso a execucao falhe!
			
	def for_connect_mysql():
		try:
			connection = msql.connect(host="localhost",db="test",   
					user="", passwd="")   
			return connection
		except Exception: 
			return ('Erro de Conexao: ',Exception) 									        # retorna um erro caso a execucao falhe!
			
	def for_query(statement,rp_dir,conexao):						   							# realiza consulta apartir do select e a rp_directory, retorna um dataframe.
		try:
			consulta = psql.read_sql_query(statement,conexao) 								# consulta apartir dos parametros informados no metodo.
			basedados = pd.DataFrame(consulta)				   								# armazena a consulta em uma base de dados onde converte o dataframe
			return basedados								   								# retorna os dados da consulta
		except Exception: 
			return ('Erro ao acessar o dataframe: ',Exception)
			
	def for_excel(directory,dbase,date,rp_dir, name_excel):	   								# realiza a insercao do dataframe no arquivo .xlsx
		try:
			print (directory)
			excel = directory+date+r'_'+name_excel.upper()+r'.xlsx'
			writer = pd.ExcelWriter(excel, engine='xlsxwriter')								# escreve o dataframe no arquivo designado anteriormente.
			dbase.to_excel(writer, date+'_'+rp_dir+'_'+name_excel,
							index = False, encoding='utf-8')								# insere todos os componentes no metodo to_excel
			writer.save()																	# Salva o progresso.
			return directory+'\\'+os.path.basename(excel) 
		except Exception: 
			return "Erro ao salvar excel: "+os.path.basename(excel)
			
	def for_zip(directory,date,rp_dir, name_db,dir_excel):									# realiza a insercao dos arquivos .xlsx em um arquivo zip
		try:
			rar = open (directory+date+'_'+rp_dir+'_'+name_db+'.zip', 'wb',encoding='UTF8')					# cria e nomeia o .rar passando os paramentos do metodo
			doc_zip = zipfile.ZipFile(rar, mode="w")										# cria e nomeia o .zip passando o rar designado anteriormente
			for file in glob.glob(directory+'\*.xlsx'):										# for para insercao do .xlsx no arquivo zipado
				doc_zip.write (os.path.join(file),
				os.path.relpath(os.path.join(file),
				directory),
				compress_type = zipfile.ZIP_DEFLATED)
			# doc_zip.close()
			# rar.close()
		except Exception:
			print ("Erro ao salvar arquivo .zip: "+os.path.basename(doc_zip))
			
	def delete_excel(directory):
		try:
			for file_excel in glob.glob(directory+'\*.xlsx'):								# for para insercao do .xlsx no arquivo zipado
				os.remove(file_excel)
		except Exception:
			return ("Erro ao deletar arquivos excel.")
			
	def create_directory(directory, ano):
		try:
			new_dir = directory + '\\'+str(ano)+'_'+datetime.date.today().strftime("%B").upper()
			if (os.path.isdir(new_dir) == True):	
				print ("Diretorio ja existe")
				return new_dir+'\\'
			else:
				os.mkdir(new_dir)
				return new_dir+'\\'
		except Exception:
			return ("Erro ao criar novo diretorio.")	
			
	if __name__ == "__main__":
		try: 
			num = 0
			directory_list = list()														# declara o array para armazenar os diretorios
			date = time.strftime("%Y%m")												# data atual do sistema 
			dir_repository = r'C:\Users\WESLEI XIMENES\Documents'
			name_db = ['name_database']													# array de nomes para os arquivos
			for root, dirs, files in os.walk(dir_repository):							# loop para encontrar diretorios e subdiretorios com a nomenclatura de cau_rp_directory
				for name_dir in dirs:													# loop para percorrer dentro do diretorio e apos verificar se existe,
					if 'directory_' in name_dir:												# armazena os diretorios em um array. 
						directory_list.append(os.path.join(root, name_dir))
			for directory in directory_list:												# percorre o diretorios armazenados e chama os metodos designados
				print (os.path.basename(directory))
				rp_directory = os.path.basename(directory)
				dir_directory = directory+r'\\'
				dir_directory = create_directory(dir_directory,time.strftime("%Y"))
				delete_excel(dir_directory)
				if 'directory' in os.path.basename(directory):												
					for name in name_db:
						sql = ( "SELECT * FROM dono_ft")
						dadosbase = for_query(sql,rp_directory[0].upper(),for_connect_mysql())
						directory_excel = for_excel(dir_directory,dadosbase,date,rp_directory,name)
						# for_zip(dir_directory,date,rp_directory[0].upper(),name_db[0],directory_excel)
						# delete_excel(dir_directory)

			#print("\nSucesso!")
		except Exception:
			print ("Erro na Classe Principal: ",Exception)