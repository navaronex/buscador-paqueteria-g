import base64
import os 
from docx import Document #Libreria docx que "leerá" el documento, Document es una Clase
import pandas as pd
import streamlit as st
from PIL import Image
import re
import unicodedata
import hashlib
import time



# ------------------------------------------- BACKEND ---------------------------------------------------------------------------------------
def leerDocumentosPorPaises(rutaDelArchivo):
    doc = Document(rutaDelArchivo)
    infoPaises = {}
    paisActual = None
    bufferDeTexto = []

    seccionesEspeciales = [
        "Países de la Unión Europea", "Envío de baterías de forma segura", "Envío de perfumes sprays Ecuador", 
        "Gastos por movilización a la aduana", "Otros artículos a exportar", "Seguro opcional", 
        "Medidas plancha mercancía ecuador", "Envío Transporte de mercancía peligrosa DGD",
        "Envío y forma de mercancía pesada y volumétrica", "Productos del sistema de paquetería", 
        "Acceso sistema PUDELECO"
    ]

    for x in doc.paragraphs:
        texto = x.text.strip()
        if not texto:
            continue
        
        buscarTexto = texto.lower()

        # 1. DETECCIÓN DE TÍTULOS (Prioridad Absoluta)
        esUE = ("unión europea" in buscarTexto or "union europea" in buscarTexto) and len(texto) < 40
        
        seccionEncontrada = None
        for s in seccionesEspeciales:
            if s.lower() in buscarTexto and len(texto) < 80:
                seccionEncontrada = s
                break

        # Filtro para países normales
        esExcepcion = (buscarTexto == "condiciones de envío:" or 
                       buscarTexto == "condiciones de envio:" or
                       (len(texto) < 25 and "condiciones" in buscarTexto))

        esPaisNormal = (not esExcepcion and not seccionEncontrada and not esUE and
                        ("[" not in texto) and 
                        ("paquetería" in buscarTexto or "paqueteria" in buscarTexto or 
                         "courier" in buscarTexto or "condiciones de" in buscarTexto) 
                        and len(texto) < 85)

        # 2. SI DETECTAMOS UN TÍTULO NUEVO, GUARDAMOS EL ANTERIOR
        if esUE or seccionEncontrada or esPaisNormal:
            if paisActual and bufferDeTexto:
                contenido = "\n".join(bufferDeTexto)
                if paisActual in infoPaises:
                    infoPaises[paisActual] += "\n" + contenido
                else:
                    infoPaises[paisActual] = contenido
                bufferDeTexto = [] # Vaciamos el buffer para el nuevo país

            # ASIGNAR EL NUEVO NOMBRE
            if esUE:
                paisActual = "Países de la Unión Europea"
            elif seccionEncontrada:
                paisActual = seccionEncontrada
            else:
                nombre = texto
                for basura in ["Categorías de paquetes", "Condiciones de", "Paquetería", "Courier"]:
                    nombre = nombre.replace(basura, "").replace(basura.lower(), "")
                 
                nombre_limpio = nombre.split("(")[0].replace(":", "").strip()

                if "españa" in nombre_limpio.lower() and "uu" in nombre_limpio.lower():
                    paisActual = "Condiciones España-EE.UU."
                else:
                    paisActual = nombre_limpio.title().strip()
        
        elif paisActual:
            # Si no es un título, es contenido y va al buffer
            bufferDeTexto.append(texto)

    # Guardar el último al salir del bucle
    if paisActual and bufferDeTexto:
        contenido = "\n".join(bufferDeTexto)
        if paisActual in infoPaises:
            infoPaises[paisActual] += "\n" + contenido
        else:
            infoPaises[paisActual] = contenido
    
    directorioScript = os.path.dirname(os.path.abspath(__file__))
    for x in infoPaises:
        lineas = infoPaises[x].split("\n")
        infoPaises[x] = formatearTexto(lineas, directorioScript)

    return infoPaises    
    

def formatearTexto(lineas, directorioBase):
    resultado = []
    palabrasTexto = [
        "Peso", "Precio","Destinatario", "Importante", "Prohibido", "Documentación", "Observaciones","Categoría", "Seguro", 
        "Medidas", "Nota", "Valor", "Declaración", "Máximo", "Mínimo", "Entregas", "Condiciones","Condiciones generales", "Decreto",
         "Descripción", "Impuesto", "Efectivo", "Menaje de casa", "PERFUMES SPRAYS", "Envío", "lunes", "martes", "miércoles", "jueves", 
         "viernes", "sábado", "domingo", "Dimensiones", "0","1","2","3","4","5","6","7","8","9", "embalaje", "reducido", "no podrá enviarse mediante el servicio de Courier",
         "cotizar el envío individual del artículo directamente con la aerolínea", "Carga volumétrica",
         "Fórmula general", "Largo", "Ancho", "Alto", "kg", "cm", "Nº de bultos", "Factura comercial"
         "CATEGORÍA G", "CATEGORÍA B"
    ]

    for x in lineas:
        l = x.strip()
        if not l:
            continue
        # ------------------------- Detectar y convertir URL's ---------------------------
        url_pattern = r'(https?://[^\s]+)'

        if "youtube.com/watch?v=" in l or "youtu.be/" in l:
            # Si es YouTube, extraemos el ID para embeber el video
            video_id = l.split("v=")[-1].split("&")[0] if "v=" in l else l.split("/")[-1]
            resultado.append(f'''
                <div style="margin: 20px 0; text-align: center;">
                    <iframe width="100%" height="315" src="https://www.youtube.com/embed/{video_id}" 
                    frameborder="0" allowfullscreen style="border-radius:10px;"></iframe>
                </div>
            ''')
            continue # Saltamos para no poner el texto de la URL debajo
        
        # Para otras URLs normales, las hacemos clicables
        l = re.sub(url_pattern, r'<a href="\1" target="_blank" style="color: #C0392B; text-decoration: underline;">\1</a>', l)
        
        # ------------------------------ FIN DETECCIÓN URLS ----------------------------

        l_low = l.lower().replace(" ", "").strip()
        # asegurar que limpie posibles puntos o caracteres del Word:
   

        # --- GESTIÓN DE IMÁGENES ROBUSTA ---
        # l_low ya viene limpio (sin espacios y en minúsculas)
        
        # Diccionario de mapeo: Etiqueta en Word -> Nombre de archivo real
        mapeo_imagenes = {
            "[imagentabla1]": "tabla1.png",
            "[imagentabla2]": "tabla2.png",
            "[imagentabla3]": "tabla3.png",
            "[imagenseguro1]": "seguro1.png",
            "[imagenseguro2]": "seguro2.png",
            "[imagenbaterias]": "baterias.png",
            "[imagencargavol]": "CargaVolumen.png",
            "[imagenformulavol]": "formulaVolumen.png",
            "[imagenformulapesovol]": "formulaPesoVol.png",
            "[imagenplancha]": "plancha.png",
            "[imagencourierec]": "courierEc.png"
        }

        l_limpia = l.lower().replace(" ", "")

        encontrada = False
        for etiqueta, archivo in mapeo_imagenes.items():
            if etiqueta in l_limpia:
                ruta_img = os.path.join(directorioBase, "assets", archivo)
                s = obtenerImagenBase64(ruta_img)
                if s:
                    resultado.append(f'<div style="text-align:center;margin:20px 0;"><img src="{s}" style="max-width:100%;border-radius:5px;border:1px solid #ccc;"></div>')
                else:
                    # ESTO ES PARA DEPURAR: Si no carga, te dirá la ruta que falló en el HTML
                    resultado.append(f'<p style="color:orange; font-size:0.7rem;">Error: No se encontró {archivo} en {ruta_img}</p>')
                encontrada = True
                break
        
        if encontrada: continue

        # Caso especial Perfumes (Doble imagen)
        if "[perfumespray1]" in l_low:
            r1 = os.path.join(directorioBase, "assets", "perfumeSpray1.png")
            r2 = os.path.join(directorioBase, "assets", "perfumeSpray2.png")
            s1 = obtenerImagenBase64(r1)
            s2 = obtenerImagenBase64(r2)
            if s1 and s2:
                resultado.append(f'''
                    <div style="display:flex;justify-content:center;gap:10px;margin:20px 0;">
                        <img src="{s1}" style="width:48%;border:1px solid #ccc;">
                        <img src="{s2}" style="width:48%;border:1px solid #ccc;">
                    </div>''')
            continue
        elif "[perfumespray2]" in l_low:
            continue

        # (El resto de tu lógica de títulos y párrafos se mantiene igual debajo)
        esPalabraClave = any(l.startswith(pc) for pc in palabrasTexto)
        esTituloCorto = (len(l) < 50 and l.endswith(":"))
        esCategoria = l.upper().startswith("CATEGORÍA") and len(l) < 40

        if esPalabraClave or esTituloCorto or esCategoria:
            if resultado:
                resultado.append('<hr style="border:0;border-top:1px solid #eee;margin:20px 0;">')
            texto = l.replace(":", "").strip()
            transformacion = "text-transform:uppercase;" if len(texto) < 30 else ""
            resultado.append(f'<p style="color:#C0392B;font-weight:bold;margin-bottom:8px;{transformacion}font-size:0.95rem;">{texto}</p>')
        else:
            texto_parrafo = l
            for pc in palabrasTexto:
                if pc in texto_parrafo: texto_parrafo = texto_parrafo.replace(pc, f"<b>{pc}</b>")
            
            if l.startswith("-") or l.startswith("*") or l.startswith("•"):
                resultado.append(f'<div style="margin-left:20px;margin-bottom:5px;color:#444;">• {texto_parrafo[1:].strip()}</div>')
            else:
                resultado.append(f'<p style="margin-bottom:10px;color:#333;text-align:justify;line-height:1.5;">{texto_parrafo}</p>')
            
    return "".join(resultado)


def estilosBuscador (archivo_css):
    if os.path.exists(archivo_css):
        with open(archivo_css) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    else:
        st.error(f"No se encontró el archivo de estilos: {archivo_css}")


def obtenerImagenBase64(rutaImagen): #Función que convertirá el logo de geomil en base64
    if os.path.exists(rutaImagen):
        extension = os.path.splitext(rutaImagen)[1].lower().replace(".", "")

        mapeoTipo = "jpeg" if extension in ["jpg", "jpeg"] else "png"

        with open(rutaImagen, "rb") as archivoDeImagen:
            stringCodificado = base64.b64encode(archivoDeImagen.read()).decode()
            return f"data:image/{mapeoTipo};base64,{stringCodificado}"
    return None


def cargarJs(nombreJs):
    ruta = os.path.join(os.path.dirname(__file__), nombreJs)
    if ruta and os.path.exists(ruta):
        with open(ruta, "r", encoding='utf-8') as f:
            return f.read()
    return ""




# ------------------------------------------------------ FRONTEND -------------------------------------------------------------------
import requests

st.set_page_config(
    page_title = "Buscador Paquetería",
    page_icon = "",
    menu_items = {
        'Get help': 'https://docs.streamlit.io/library/get-started/quick-start',
        'Report a bug': 'https://www.senescyt.gob.ec/web/guest/consultas',
        'About': 'http://tucelularlegal.arcotel.gob.ec/tucelularlegal/consulta_Imeis.aspx'
    }
) #1º
directorioActual = os.path.dirname(__file__) #2º

# ----------------- Configuración para la Nube de Streamlit -----------------
idArchivoDocx = "11VPTTWBCUAityL5okT2sAm7JM8eNc-DT"
rutaReal = os.path.join(directorioActual, "condicionesPaqueteria.docx") 

@st.cache_data(ttl = 600) #Actualizar (cada 10 min)
def descargarWord(idArchivo, rutaDestino):
    url = f'https://docs.google.com/document/d/{idArchivo}/export?format=docx'
    try:
        respuesta = requests.get(url, timeout=15)
        if respuesta.status_code == 200:
            with open (rutaDestino, 'wb') as f:
                f.write(respuesta.content)
    except:
        pass

descargarWord (idArchivoDocx, rutaReal)


ruta_css = os.path.join(directorioActual, "estilos.css") #Aplicar CSS al sitio web (3º)
estilosBuscador(ruta_css)


########################################## IMAGEN GEOMIL ############################################
rutaLogo = os.path.join(directorioActual, "assets", "LogoGeomilBlanco.png") # 4º

rutaB1 = os.path.join(directorioActual, "assets", "logoArancel.png")
rutaB2 = os.path.join(directorioActual, "assets", "logoSenecsyt.png")
rutaB3 = os.path.join(directorioActual, "assets", "logoTuCelular.png")


rutaHTML = os.path.join(directorioActual, "header.html")

if os.path.exists(rutaLogo) and os.path.exists(rutaHTML):
    logoBase64 = obtenerImagenBase64(rutaLogo) #Convertir el logo a imagen base64

    boton1 = obtenerImagenBase64(rutaB1)
    boton2 = obtenerImagenBase64(rutaB2)
    boton3 = obtenerImagenBase64(rutaB3)

    
    with open(rutaHTML, "r", encoding="utf-8") as f:
        st.markdown (f.read().replace("{{LOGO_BASE64}}", logoBase64), unsafe_allow_html=True)

        st.markdown(f"""
        <style>
            /* Contenedor para los 3 botones arriba */
            .barra-botones-top {{
                position: fixed;
                top: 8px;
                right: 310px; /* Separación del botón Deploy/Menú */
                z-index: 999999;
                display: flex;
                gap: 20px;
                align-items: center;
            }}
            .img-top {{
                width: 50px;
                height: 40px;
                object-fit: contain;
                cursor: pointer;
                transition: 0.3s;
                filter: drop-shadow(0px 0px 3px rgba(255,255,255,0.8));
            }}
            .img-top:hover {{ 
                transform: scale(1.2); 
                filter: drop-shadow(0px 0px 6px rgba(255,255,255,0.8));
            }}
            
            /* Ajuste para móviles: si la pantalla es pequeña, los ocultamos o movemos */
            @media (max-width: 800px) {{ .barra-botones-top {{ display: none; }} }}
        </style>
        
        <div class="barra-botones-top">
            <a href="https://www.pudeleco.com/online/clave.html" target="_blank">
                <img src="{boton1}" class="img-top" title="Aranceles Pudeleco">
            </a>
            <a href="https://www.senescyt.gob.ec/web/guest/consultas" target="_blank">
                <img src="{boton2}" class="img-top" title="Senescyt Ecuador">
            </a>
            <a href="http://tucelularlegal.arcotel.gob.ec/tucelularlegal/consulta_Imeis.aspx" target="_blank">
                <img src="{boton3}" class="img-top" title="Tucelularlegal">
            </a>
        </div>
        """, unsafe_allow_html=True)



#####################################################################################################


if os.path.exists(rutaReal): # Si el archivo existe en esa ruta.... ES para comprobar que está el docx
    datos = leerDocumentosPorPaises(rutaReal)

    listaPaises = list(datos.keys()) #Ordenar el desplegable, 1º Courier
    if "Ecuador" in listaPaises: #SI está ecuador
        listaPaises.remove("Ecuador") #Remover ecuador...
        listaPaises.insert(0, "Condiciones menaje de casa") #Insertar en la primera posición (al inicio)
        datos["Condiciones menaje de casa"] = datos.pop("Ecuador") #Actualizar el "diccionario"


    col1, col2, col3 = st.columns([1, 2, 1]) #Centrar el buscador en columnas... Selector de secciones
    with col2:
        # Aquí 'orden' garantiza que el selectbox mantenga el orden del Word
        opcion = st.selectbox("Selecciona sección:", [""] + listaPaises) 

    if opcion:
        st.markdown(f'<p class = "titulo-centrado-rojo">{opcion}</p>', unsafe_allow_html = True) #Título pais elegido
        
        with st.sidebar: # Buscador en la barra lateral
            st.header("Buscador")
            buscador = st.text_input ("Escribe para buscar:", placeholder = "Factura, 100kg,...", key = "input_busqueda")

            contenidoHTML = datos [opcion]

        ####################### Buscador ####################º
            if buscador:
                def generarPatronFlexible(palabra):
                    dic = { 'a': '[aá]',
                        'e': '[eé]',
                        'i': '[ií]',
                        'o': '[oó]',
                        'u': '[uú]',
                        }
                    return "".join(dic.get(c.lower(), re.escape(c)) for c in palabra)
            
                patronStr = generarPatronFlexible(buscador)
                patron = re.compile(patronStr, re.IGNORECASE)

                matches = list(patron.finditer(contenidoHTML))
                coincidencias = len(matches)

                if coincidencias > 0:
                #st.info(f"Se ha encontrado {coincidencias} coincidencias.")
                    if 'indiceBusqueda' not in st.session_state:
                        st.session_state.indiceBusqueda = 1

                    if 'ultimaPalabra' not in st.session_state or st.session_state.ultimaPalabra != buscador:
                        st.session_state.indiceBusqueda = 1
                        st.session_state.ultimaPalabra = buscador
                    st.write(f"Resultados: **{coincidencias}**")

                    col_prev, col_num, col_next = st.columns([1, 2, 1])
                
                    def izq():
                        st.session_state.indiceBusqueda = max (1, st.session_state.indiceBusqueda - 1)

                    def drcha():
                        st.session_state.indiceBusqueda = min (coincidencias, st.session_state.indiceBusqueda + 1)

                    with col_prev: 
                        st.button("<-", key = "prev_side", on_click = izq)
                          
                    with col_num:
                        st.number_input (
                            label ="Navegador",
                            min_value = 1,
                            max_value  = coincidencias,
                            key = "indiceBusqueda",
                            label_visibility = "collapsed",
                            step = 1
                        )

                    with col_next:
                        st.button("->", key = "next_side", on_click = drcha)

                    st.markdown(f"<div style= 'text-align: center; font-size: 0.8rem; color: #666; margin-top: -10px;'>{st.session_state.indiceBusqueda} de {coincidencias} </div>", unsafe_allow_html = True)
                    indiceActual = st.session_state.indiceBusqueda


                #Subrayado de palabras del buscador
                    listaContador = {"n": 0}
                    def sustituir(match):
                        listaContador["n"] += 1
                        texto = match.group(0)

                        matchId = f"match_{listaContador['n']}"

                        clase = "subrayado-busqueda active" if listaContador ["n"] == indiceActual else "subrayado-busqueda"

                        return f'<span id="{matchId}" class="{clase}">{texto}</span>'

                    contenidoHTML = patron.sub(sustituir, contenidoHTML)
                
                

                        
                    codigoJs = cargarJs("buscador.js")
                    if codigoJs:
                        idUnico = int (time.time()*1000)
                        
                        scriptFinal = f"<script> window.targetMatchId = 'match_{indiceActual}'; {codigoJs} // ID de ejecución: {idUnico}     </script>"                    
                        st.components.v1.html (scriptFinal, height = 0)

        ########### Imagen de Courier Ecuador ################
        #if opcion == "Ecuador":
        #    rutaImgEc = os.path.join(directorioActual, "assets", "courierEc.png")
        #    if os.path.exists(rutaImgEc):
        #        st.image(rutaImgEc, use_container_width=True)
        ########### Imagen de Baterías #########################

        

        st.markdown(f'<div class = "resultado-caja">{contenidoHTML}</div>', unsafe_allow_html = True)
