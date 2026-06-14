\# Asistente Científico para Elaboración de Artículos Académicos



Aplicación desarrollada en \*\*Streamlit\*\* para apoyar la elaboración de artículos científicos a partir de fuentes académicas en PDF, metadatos bibliográficos de Zotero, matrices de revisión documental, redes bibliométricas compatibles con VOSviewer y herramientas de revisión académica asistidas por inteligencia artificial.



\## Descripción general



Esta aplicación permite construir borradores de artículos académicos de manera estructurada, utilizando fuentes documentales reales. Su propósito es apoyar procesos de revisión bibliográfica, análisis de literatura científica, organización de referencias, generación de artículos por secciones y evaluación de calidad académica.



La app no reemplaza la revisión del autor, investigador o comité académico. El contenido generado debe ser validado antes de enviarse a una revista, congreso o repositorio institucional.



\## Funcionalidades principales



\* Carga manual de artículos científicos en PDF.

\* Extracción de texto desde documentos PDF.

\* Conexión con Zotero para importar metadatos bibliográficos.

\* Intento de descarga y lectura de PDF adjuntos desde Zotero.

\* Generación de matriz bibliográfica.

\* Organización de referencias en estilos académicos.

\* Generación de redes bibliométricas compatibles con VOSviewer.

\* Creación de red de palabras clave.

\* Creación de red de coautoría.

\* Generación de artículos completos.

\* Generación de artículos por secciones.

\* Selección de plantillas académicas.

\* Revisión de calidad académica mediante rúbrica.

\* Revisión de coherencia interna del manuscrito.

\* Generación de resumen estructurado.

\* Generación de abstract y keywords en inglés.

\* Paráfrasis académica.

\* Exportación en Word, PDF, TXT y Excel.

\* Guardado y carga de proyectos locales.



\## Estructura del proyecto



```text

Elaboracion\_articulos/

│

├── aplicacion.py

├── requirements.txt

├── README.md

├── .gitignore

│

├── .streamlit/

│   └── secrets.toml

│

└── proyectos\_guardados/

```



\## Instalación



Clone o descargue este repositorio y abra una terminal en la carpeta del proyecto.



Instale las dependencias con:



```bash

pip install -r requirements.txt

```



\## Dependencias



El archivo `requirements.txt` debe incluir:



```txt

streamlit

openai

pypdf

fpdf

python-docx

pandas

openpyxl

requests

```



\## Configuración de claves



La aplicación requiere una clave API de OpenAI.



Cree una carpeta llamada `.streamlit` y dentro de ella un archivo llamado `secrets.toml`.



La estructura debe quedar así:



```text

.streamlit/

└── secrets.toml

```



Dentro de `secrets.toml`, agregue:



```toml

OPENAI\_API\_KEY = "SU\_CLAVE\_DE\_OPENAI"

```



No suba este archivo a GitHub.



\## Archivo `.gitignore`



Para proteger las claves y archivos temporales, el proyecto debe incluir un archivo `.gitignore` con el siguiente contenido:



```gitignore

.streamlit/secrets.toml

\_\_pycache\_\_/

\*.pyc

proyectos\_guardados/

.env

```



\## Ejecución local



Para ejecutar la aplicación:



```bash

streamlit run aplicacion.py

```



Luego abra en el navegador la dirección que indique Streamlit, normalmente:



```text

http://localhost:8501

```



\## Flujo de uso recomendado



1\. Cargar artículos científicos en PDF o importar registros desde Zotero.

2\. Extraer el texto de las fuentes.

3\. Generar la matriz bibliográfica.

4\. Organizar referencias.

5\. Generar redes bibliométricas para VOSviewer.

6\. Completar título, tema, objetivo y metodología.

7\. Seleccionar una plantilla académica.

8\. Generar el artículo por secciones.

9\. Revisar calidad académica.

10\. Revisar coherencia interna.

11\. Generar resumen, abstract y palabras clave.

12\. Exportar el resultado en Word, PDF, TXT o Excel.

13\. Guardar el proyecto para continuar posteriormente.



\## Integración con Zotero



La aplicación permite importar metadatos bibliográficos desde Zotero mediante:



\* User ID o Group ID.

\* API Key de Zotero.

\* Selección de colecciones.

\* Importación de títulos, autores, años, revistas, DOI, URL, etiquetas y resúmenes.

\* Intento de lectura de PDF adjuntos, cuando los permisos de Zotero lo permiten.



\## Integración con VOSviewer



La aplicación genera archivos compatibles con VOSviewer para análisis bibliométrico:



```text

map\_keywords.txt

network\_keywords.txt

map\_authors.txt

network\_authors.txt

```



Estos archivos permiten construir redes de:



\* Coocurrencia de palabras clave.

\* Coautoría.



\## Exportaciones disponibles



La aplicación permite descargar:



\* Artículo en Word `.docx`.

\* Artículo en PDF `.pdf`.

\* Artículo en texto plano `.txt`.

\* Matriz bibliográfica en Excel `.xlsx`.

\* Archivos de red para VOSviewer `.txt`.



\## Advertencia académica



Esta herramienta es un apoyo para procesos de escritura científica. El usuario debe revisar, validar y ajustar el contenido generado. La aplicación no garantiza aceptación en revistas indexadas ni reemplaza el trabajo investigativo, metodológico o editorial del autor.



Se recomienda verificar:



\* Originalidad del texto.

\* Validez de las referencias.

\* Coherencia metodológica.

\* Cumplimiento de normas de la revista.

\* Revisión antiplagio.

\* Revisión por pares o asesor académico.



\## Autor



Proyecto desarrollado como herramienta de apoyo para la elaboración de artículos académicos y científicos.



\## Licencia



Este proyecto puede ser adaptado para fines académicos, investigativos y formativos.



