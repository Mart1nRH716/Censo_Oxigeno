import os
import pandas as pd
import pyodbc
from pyxlsb import open_workbook 



# -------------------------------
# CONFIGURACIÓN
# -------------------------------
CARPETA_REPORTES = "2024_oxigeno"  

CONEXION_SQL = (
	"DRIVER={SQL Server};"
	# "SERVER=10.250.11.237\DB_CIS_IMSS;"  
	"SERVER=localhost;"
	"DATABASE=BD_OXIGENO;"  # Las pruebas se realizaron en local para despues migrar las tablas al 237
	#Poner usuario y contraseña en caso de ser necesario, si se deja comentado, este tomará la autenticación de windows
	# "UID=martin.romeroh;" # ------------------Poner usuario  y contraseña              
	# "PWD=ConsolaAzul710;"             
)
TABLA_DESTINO = "censo_oxigeno"


estados_dict = {
	"AGS": "Aguascalientes",
	"BCN": "Baja California",
	"BCS": "Baja California Sur",
	"CAMP": "Campeche",
	"CHIAP": "Chiapas",
	"CHIHUA": "Chihuahua",
	"COAH": "Coahuila",
	"COL": "Colima",
	"DF NORTE": "CDMX Norte",
	"DF SUR": "CDMX Sur",
	"DGO": "Durango",
	"EDO DE MEX OTE": "Estado de México Oriente",
	"EDO DE MEX PTE": "Estado de México Poniente",
	"GRO": "Guerrero",
	"GTO": "Guanajuato",
	"HGO": "Hidalgo",
	"JAL": "Jalisco",
	"MICH": "Michoacán",
	"MOR": "Morelos",
	"NAY": "Nayarit",
	"NL": "Nuevo León",
	"OAX": "Oaxaca",
	"PUE": "Puebla",
	"QROO": "Quintana Roo",
	"QUER": "Querétaro",
	"SIN": "Sinaloa",
	"SLP": "San Luis Potosí",
	"SON": "Sonora",
	"TAB": "Tabasco",
	"TAM": "Tamaulipas",
	"TLAX": "Tlaxcala",
	"VER NTE": "Veracruz (Norte)",
	"VER SUR": "Veracruz (Sur)",
	"YUC": "Yucatán",
	"ZAC": "Zacatecas"
}



def leer_datos(ruta_archivo):
	try:
		
		
		df = pd.read_excel(ruta_archivo, engine="pyxlsb", header=2)
		df = df.drop(columns=[col for col in df.columns if str(col).replace('.', '').isdigit()], errors="ignore")
		df.rename(columns={col : col.lower().replace(' ', '_') for col in df.columns}, inplace=True)
		df = df.iloc[:-1]
		if ruta_archivo.endswith(".xlsb"):
			extension = ".xlsb"
		else:
			extension = ".XLSB"
		archivo = os.path.basename(ruta_archivo).replace(extension, "").upper()
		print(archivo)
		df["delegacion"] = estados_dict.get(archivo) 
		df["anio"] = CARPETA_REPORTES.split("_")[0]

		columnas_fecha = ["fecha_inicio", "fecha_finaliza_receta", "fecha_nacimiento"]
		columnas_varchar = ["folio_receta", "nss", "agregado", "telefono", "celular", "correoe", "delegacion",
			"unidad_adscrip", "desc_corta_adscrip", "tipo_unidad_adscrip", "medico_matricula"]
		columnas_float = [
			"estatura", "peso", "flujo", "oxigeno_iva", "cpab/bpap",
			"precio", "precio_c/iva", "total", "cobro_iva"
		]
		columnas_int = [
			"dias_oxigeno", "periodo", "contador",
			"tanque_oxigeno", "tanque_portatil", "concentrador",
			"cpap", "bpap", "nebulizador", "anio"
		]
  
		#Fechas
		# for col in columnas_fecha:
		# 	if col in df.columns:
		# 		#C on esta linea se subieron los archivos, sin embargo al moemento de cargar OAXACA, no cargaba el archivo. (Revisar si en las demás este cambio no afecta)
		# 		df[col] = pd.to_datetime(df[col], errors="coerce", unit="d", origin="1899-12-30") 
		# 		# if archivo != 'OAX':
		# 		# 	df[col] = pd.to_datetime(df[col], errors="coerce", unit="d", origin="1899-12-30") 
		# 		# else:
		# 		# 	df[col] = pd.to_datetime(df[col], errors="coerce")

		for col in columnas_fecha:
			if col in df.columns:
				# print(f"Columna {col} - tipo: {df[col].dtype}")
				# print(f"Columna {col} - primeros valores: {df[col].head(3).tolist()}")
				
				if df[col].dtype in ['float64', 'int64']:
					# Convertir desde números de serie de Excel
					df[col] = pd.to_datetime(df[col], errors="coerce", unit='D', origin='1899-12-30')
				else:
					# Intentar conversión directa
					df[col] = pd.to_datetime(df[col], errors="coerce")
        

		
		# Floats
		for col in columnas_float:
			if col in df.columns:
				df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(float)

		# Ints
		for col in columnas_int:
			if col in df.columns:
				df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    
		for col in columnas_varchar:
			if col in df.columns:
				if col == 'medico_matricula' or col == 'unidad_adscrip':
					# Convertir a string sin notación científica y sin .0
					df[col] = df[col].apply(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x))
				else:
					df[col] = df[col].astype(str)
					df[col] = df[col].str.lstrip("'")
		
		# for col in df.columns:
		# 	df[col] = df.fillna(" ")
		return df
		
	except Exception as e:
		print(f"Error leyendo {ruta_archivo}: {str(e)}")
		return pd.DataFrame()  # devuelve df vacío si falla


def insertar_en_sql(registros, conexion_str, tabla_destino):
	try:
		#Inserción de los regitros con un cursor
		conn = pyodbc.connect(conexion_str)
		cursor = conn.cursor()
		# Creamos la tabla si no existe
		cursor.execute(f"""
		IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{tabla_destino}' AND xtype='U')
		CREATE TABLE {tabla_destino}(
			id INT IDENTITY(1,1) PRIMARY KEY,
			anio int,
			delegacion VARCHAR(100),
			folio_receta VARCHAR(50) UNIQUE,
			fecha_inicio DATE,
			fecha_finaliza_receta DATE,
			dias_oxigeno INT,
			nss VARCHAR(20),
			agregado VARCHAR(20),
			nombre_paciente VARCHAR(150),
			fecha_nacimiento DATE,
			estado_civil VARCHAR(50),
			ocupacion VARCHAR(100),
			estatura FLOAT,
			peso FLOAT,
			calle VARCHAR(250),
			numero VARCHAR(20),
			interior VARCHAR(20),
			cruce1 VARCHAR(250),
			cruce2 VARCHAR(250),
			colonia VARCHAR(350),
			codigo_postal VARCHAR(10),
			referencia_domicilio VARCHAR(350),
			telefono VARCHAR(50),
			celular VARCHAR(50),
			correoe VARCHAR(100),
			deleg_adscrip VARCHAR(50),
			unidad_adscrip VARCHAR(50),
			desc_corta_adscrip VARCHAR(100),
			tipo_unidad_adscrip VARCHAR(100),
			deleg_expide VARCHAR(10),
			unidad_expide VARCHAR(10),
			desc_corta_expide VARCHAR(30),
			tipo_unidad_expide VARCHAR(30),
			tanque_oxigeno BIT,
			tanque_portatil BIT,
			concentrador BIT,
			cpap BIT,
			bpap BIT,
			nebulizador BIT,
			flujo FLOAT,
			periodo int,
			diagnostico VARCHAR(50),
			descripcion_diagnostico VARCHAR(MAX),
			medico_matricula VARCHAR(50),
			nombre_medico VARCHAR(150),
			oxigeno_iva FLOAT,
			[cpab/bpap] FLOAT, 
			precio FLOAT,
			[precio_c/iva] FLOAT,
			cobro_iva FLOAT,
			contador INT,
			total FLOAT,
		) ON [PRIMARY]
		""")

		contador = 0
		for idx, reg in registros.iterrows():
			try:
				# Filtrar solo columnas con valores no nulos
				datos_validos = {col: val for col, val in reg.items() if pd.notnull(val)}
				if not datos_validos:
					continue

				columnas = ", ".join(f"[{col}]" for col in datos_validos.keys())
				placeholders = ", ".join("?" for _ in datos_validos)
				valores = list(datos_validos.values())

				query = f"INSERT INTO {tabla_destino} ({columnas}) VALUES ({placeholders})"
				cursor.execute(query, valores)
				contador += 1
			except Exception as e:
				print(f"Error en fila {idx}: {e}")
				#print(f"Valores problemáticos: {datos_validos}")
				continue  # continuar con el siguiente registro si hay error

		conn.commit()
		conn.close()
		print(f"Se insertaron {contador} registros de {len(registros)}.")
	except Exception as e:
		print(f"Error al insertar en SQL: {str(e)}")


def procesar_carpeta(carpeta):
	# Obtener lista de archivos .xlsx en la carpeta para insertar todos los reportes dentro de la db
	archivos = [f for f in os.listdir(carpeta) if f.endswith('XLSB') or f.endswith('xlsb')]
	
	if not archivos:
		print(f"No se encontraron archivos .XLSB en la carpeta {carpeta}")
		return
	
	# archivos = ['AGS.XLSB', 'OAX.XLSB' ]
	# archivos = ['OAX.XLSB' ]
	print(archivos)
	
	total_registros = 0
	
	for archivo in archivos:
		ruta_completa = os.path.join(carpeta, archivo)
		print(f"\nProcesando archivo: {archivo}")
		
		
		datos = leer_datos(ruta_completa)
		insertar_en_sql(datos, CONEXION_SQL, TABLA_DESTINO)
		total_registros += len(datos)
		

	print(f"\nProceso completado. Total de registros insertados: {total_registros}")

procesar_carpeta(CARPETA_REPORTES)