## Automatización de la carga de los reportes de excel del Censo de Oxígeno

### Características principales
* **Automatiza la carga de los archivos de excel del censo de Oxígeno en la base de datos.
* **Almacenamiento de archivos:** el programa realiza una búsqueda en las carpetas de "{AÑO}_oxigeno" Para obtener los archivos sin la necesidad de agregarlos al código.

### Requisitos del sistema
* **Pýthon**  >= 3.8
* **Dependencias:** Las dependencias adicionales se listan en el archivo `requirements.txt`.
* Nota: Para poder ejecutar el script, se deben de almacenar los resportes en la siguiente ruta: "{AÑO}_oxigeno".

### Instalación
1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/Mart1nRH716/Censo_Oxigeno.git
   
## Instrucciones de Configuración
2. **Crear un entorno virtual
```bash
python -m venv venv
venv\Scripts\activate     # En Windows
source venv/bin/activate   # En Linux/Mac
```

## Instalar las dependencias
3. ** Ejecutar el siguiente comando:
```bash
pip install -r requirements.txt
```

Con estos pasos se pueden ejecutar el archivo .py
