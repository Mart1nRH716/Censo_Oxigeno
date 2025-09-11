## Automatización de la carga de los reportes de excel del Censo de Oxígeno

### Características principales
* Automatiza la carga de los archivos de excel del censo de Oxígeno en la base de datos.
* Almacenamiento de archivos:** el programa realiza una búsqueda en las carpetas de "{AÑO}_oxigeno" Para obtener los archivos sin la necesidad de agregarlos al código.

### Requisitos del sistema
* **Python**  >= 3.8
* **Dependencias:** Las dependencias adicionales se listan en el archivo `requirements.txt`.
* Nota: Para poder ejecutar el script, se deben de almacenar los resportes en la siguiente ruta: "{AÑO}_oxigeno".

### Instalación
1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/Mart1nRH716/Censo_Oxigeno.git
   
## Instrucciones de Configuración
2. Crear un entorno virtual
```bash
python -m venv venv
venv\Scripts\activate     # En Windows
source venv/bin/activate   # En Linux/Mac
```

## Instalar las dependencias
3. Ejecutar el siguiente comando:
```bash
pip install -r requirements.txt
```
## Estandarizar los arhivos .xlsb
4. Debido a que los archivos del censo de oxígeno vienen separados por OOAD, no todos siguen la misma estructura de datos. Si bien, el script maneja diferentes escenarios posibles para la lectura de los mismo, aún se necesita normalizar algunos archivos como lo es el inicio de los datos. La mayoría de los datos empiezan en la fila 3, por otro lado, no todos empiezan en dicha fila, es por ello que se necesita revisar que todos los reportes empiecen en la fila ya mencionada, teniendo cuidado de que al mover los datos no se muevan las celdas de referencia. 

Con estos pasos se pueden ejecutar el archivo .py
