import base64
import os 
from docx import Document #Libreria docx que "leerá" el documento, Document es una Clase
import pandas as pd
import streamlit as st
from PIL import Image
import base64
import re
import unicodedata
import hashlib
import time



# ------------------------------------------- BACKEND ---------------------------------------------------------------------------------------
def leerDocumentosPorPaises(rutaDelArchivo): # MÉTODO public String leerDocumentosPorPaises(String rutaDelArchivo){}
    # OBJETOS
    doc = Document(rutaDelArchivo)
    infoPaises = {}
    paisActual = None
    bufferDeTexto = [] # Buffer

    seccionesEspeciales = [
        "Países de la Unión Europea", "Enviar baterías de forma segura", "Envio de perfumes", "Gastos por movilización", "Otros artículos a exportar", "Seguro opcional", "Medidas plancha mercancía ecuador"
    ]


    for x in doc.paragraphs: #Bucle FOR para recorrer los párrafos del documento
        texto = x.text.strip() #Quitar espacios en blanco

        if not texto: #Si en el texto hay párrafos que están vacíos...
            continue
        
        buscarTexto = texto.lower() # if texto.contains("Condiciones de Paquetería"):
        # Filtro de los títulos para cada país

        esExcepcion = (buscarTexto == "condiciones de envío:" or 
                       buscarTexto == "condiciones de envio:" or
                       len(texto) < 25 and "condiciones" in buscarTexto)

        #Detector de secciones
        esUE = "unión europea" in buscarTexto or "union europea" in buscarTexto
        esSeccionEspecial = any(seccion.lower() in buscarTexto for seccion in seccionesEspeciales) and len(texto) < 75
        tituloPaises = (not esExcepcion and ("paquetería" in buscarTexto or 
                                            "paqueteria" in buscarTexto or #Detectar si es un país normal (paqueteria, courier)
                                            "courier" in buscarTexto or
                                            "condiciones de" in buscarTexto) and len(texto) < 85)


        if esUE or esSeccionEspecial or tituloPaises: #Antes de ir al siguiente titulo del país, se guarda el anterior en el BUFFER

            if paisActual and bufferDeTexto: #Si ya hay un pais actual
                contenido = "\n" .join(bufferDeTexto) ## nuevo, de la función formatearTexto
                if paisActual in infoPaises:
                    infoPaises [paisActual] += "\n"+ contenido #Une el buffer con DOS salto de linea para unir todas las lineas de texto y se guarda.
                else:
                    infoPaises[paisActual] = contenido

            if esUE:
                paisActual = "Países de la Unión Europea"
            elif esSeccionEspecial:
                for x in seccionesEspeciales:
                    if x.lower() in buscarTexto:
                        paisActual = x
                        break

            else:
                nombre = texto
                for basura in ["Categorías de paquetes", "Condiciones de", "Paquetería", "Courier"]:
                    nombre = nombre.replace(basura, "").replace(basura.lower(), "") 
                paisActual = nombre.split("(")[0].replace(":", "").strip().capitalize() #Limpieza final

            bufferDeTexto = []
        
        elif paisActual:
            bufferDeTexto.append(texto)




    if paisActual and bufferDeTexto: #Guardar último país procesado tras el bucle FOR
        contenido = "\n" .join(bufferDeTexto)
        if paisActual in infoPaises:
            infoPaises[paisActual] += "\n" + contenido
        else:
            infoPaises[paisActual] = contenido
    
    directorioScript = os.path.dirname(__file__)
    for x in infoPaises: #Formatear todo al final (para html, negrita,...)
        lineas = infoPaises[x].split("\n")
        infoPaises[x] = formatearTexto(lineas, directorioScript)


    return infoPaises    
    

def formatearTexto(lineas, directorioBase):
    resultado = []
    palabrasTexto = [
        "Peso", "Precio","Destinatario", "Importante", "Prohibido", "Documentación", "Observaciones","Categoría", "Seguro", "Medidas", "Nota", "Valor", "Declaración", "Máximo", "Mínimo", "Entregas", "Condiciones","Condiciones generales", "Decreto", "Descripción", "Impuesto", "Efectivo", "Menaje de casa", "PERFUMES SPRAYS", "Envío", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"
    ]

    rutaImgBaterias = os.path.join(directorioBase, "assets", "baterias.png")
    srcBaterias = obtenerImagenBase64(rutaImgBaterias)

    ignorar = [f"[imagen{i}]" for i in range (2,11)]
    #ignorar.append("[perfumeSpray2]") # Por si acaso el Word tiene la segunda etiqueta

    for x in lineas:
        l = x.strip()
        if not l: continue

        if l.lower() in ignorar or (l.startswith("[perfumeSpray") and l != "[perfumeSpray1]")or (l.startswith("[imagenSeguro") and l != "[imagenSeguro1]"):
            continue
        # Caso Imagen Batería
        if l == "[INSERTAR_IMAGEN_BATERIA_AQUI]":
            if srcBaterias:
                resultado.append(f'<div style="text-align:center;margin:20px 0;"><img src="{srcBaterias}" style="max-width:100%;height:auto;border:1px solid #ccc;"><p style="font-size:0.8rem;color:#666;font-style:italic;">Tabla de baterías</p></div>')
            continue


        # --- CASO PERFUMES (Simple y Directo) ---
        if l == "[perfumeSpray1]":
            src1 = obtenerImagenBase64(os.path.join(directorioBase, "assets", "perfumeSpray1.png"))
            src2 = obtenerImagenBase64(os.path.join(directorioBase, "assets", "perfumeSpray2.png"))
            
            if src1 and src2:
                resultado.append(f'''
                    <div style="display: flex; justify-content: center; gap: 10px; margin: 20px 0;">
                        <img src="{src1}" style="width: 48%; height: auto; border: 1px solid #ccc;">
                        <img src="{src2}" style="width: 48%; height: auto; border: 1px solid #ccc;">
                    </div>
                ''')
                
                continue # Solo saltamos la línea [perfumeSpray1] porque ya pusimos las dos fotos
        
        # -- IMAGENES SEGUROS --
        if l == "[imagenSeguro1]":
            srcSeg1 = obtenerImagenBase64(os.path.join(directorioBase, "assets", "imagenSeguro1.png")) 
            srcSeg2 = obtenerImagenBase64(os.path.join(directorioBase, "assets", "imagenSeguro2.png")) 

            if srcSeg1 and srcSeg2:
                htmlSeguros = f'''
                <div style="text-align: center; margin: 20px 0; padding: 10px; background-color: #f9f9f9; border-radius: 10px;">
                    <img src="{srcSeg1}" style="display: inline-block; width: 45%; max-width: 300px; margin: 5px; border: 1px solid #ddd; border-radius: 5px;">
                    <img src="{srcSeg2}" style="display: inline-block; width: 45%; max-width: 300px; margin: 5px; border: 1px solid #ddd; border-radius: 5px;">
                </div>
                '''
                resultado.append(htmlSeguros)
                continue
        
        # -- IMAGEN PLANCHA -- 
        if l == "[imagenPlancha]":
            srcPlancha = obtenerImagenBase64(os.path.join(directorioBase, "assets", "imagenPlancha.png"))
            if srcPlancha:
                resultado.append(f'''
                    <div style="text-align:center;margin:20px 0;">
                        <img src="{srcPlancha}" style="max-width:100%;height:auto;border:1px solid #ccc;">
                        <p style="font-size:0.8rem;color:#666;font-style:italic;">Medidas de plancha</p>
                    </div>
                ''')
                continue



        # Caso Galería
        etiquetaImagen = re.match(r"\[imagen\s?(\d+)\]", l.lower())
        if etiquetaImagen and int(etiquetaImagen.group(1)) == 1:
            imagenesHTML = ""
            for i in range(1, 11):
                src = obtenerImagenBase64(os.path.join(directorioBase, "assets", f"imagen{i}.jpg"))
                if not src: 
                    src = obtenerImagenBase64(os.path.join(directorioBase, "assets", f"imagen{i}.png"))
                if src:
                    imagenesHTML += f'<div style="flex:1 1 300px;max-width:400px;margin:10px;text-align:center;"><img src="{src}" style="width:100%; height:auto;"></div>'
            if imagenesHTML:
                resultado.append(f'<div style="display:flex;flex-wrap:wrap;justify-content:center;gap:10px;margin:25px 0;background:transparent;">{imagenesHTML}</div>')
            continue 

        # Formato de Títulos y Texto
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
    if ruta:
        with open(ruta, "r", encoding='utf-8') as f:
            return f.read()
        return ""




# ------------------------------------------ FRONTEND --------------------------------------------
import requests

st.set_page_config(page_title = "Buscador Paquetería", page_icon = "") #1º
directorioActual = os.path.dirname(__file__) #2º

# ----------------- Configuración para la Nube de Streamlit -----------------
idArchivoDocx = "11VPTTWBCUAityL5okT2sAm7JM8eNc-DT"
rutaReal = os.path.join(directorioActual, "condicionesPaqueteria.docx") 

@st.cache_data(ttl = 600) #Actualizar (cada 10 min)
def descargarWord(idArchivo, rutaDestino):
    url = f'https://docs.google.com/document/d/{idArchivo}/edit'
    respuesta = requests.get(url)
    if respuesta.status_code == 200:
        with open (rutaDestino, 'wb') as f:
            f.write(respuesta.content)
    else:
        st.error ("No se pudo descargar el documento...")

descargarWord (idArchivoDocx, rutaReal)


ruta_css = os.path.join(directorioActual, "estilos.css") #Aplicar CSS al sitio web (3º)
estilosBuscador(ruta_css)


########################################## IMAGEN GEOMIL ############################################
rutaLogo = os.path.join(directorioActual, "assets", "LogoGeomilBlanco.png") # 4º
rutaHTML = os.path.join(directorioActual, "header.html")


if os.path.exists(rutaLogo) and os.path.exists(rutaHTML):
    logoBase64 = obtenerImagenBase64(rutaLogo) #Convertir el logo a imagen base64

    if logoBase64:
        with open(rutaHTML, "r", encoding="utf-8") as f:
            contenidoHTML = f.read()

        finalHTML = contenidoHTML.replace("{{LOGO_BASE64}}", logoBase64)
        
        st.markdown(finalHTML, unsafe_allow_html=True)


       # st.write(f'<div class="header-container">{finalHTML}</div>', unsafe_allow_html=True)


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
        opcion = st.selectbox("Selecciona o escribe un país: ",[""]+listaPaises) #Barra de búsqueda con autocompletado

    if opcion:
        st.markdown(f'<p class = "titulo-centrado-rojo"> Información sobre {opcion}</p>', unsafe_allow_html = True) #Título pais elegido
        
        with st.sidebar: # Buscador en la barra lateral
            st.header("Buscador")
            buscador = st.text_input ("Escribe una palabra a buscar: ", placeholder = "Ej: Factura, 100kg, Moviles,...", key = "input_busqueda")

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
                #Selector de coincidencias
                #if coincidencias > 1:
                    #_, col_num, _ = st.columns([1, 2, 1])
                    #with col_num:
                        #indiceActual = st.number_input("Ir a: ", min_value=1, max_value=coincidencias, step=1)

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
        if opcion == "Condiciones menaje de casa":
            rutaImgEc = os.path.join(directorioActual, "assets", "courierEc.png")
            if os.path.exists(rutaImgEc):
                st.image(rutaImgEc, caption = "Courier Ecuador", use_container_width=True)
        ########### Imagen de Baterías #########################

        

        st.markdown(f'<div class = "resultado-caja">{contenidoHTML}</div>', unsafe_allow_html = True)




        
        




















#-- BLOQUE DE PRUEBA... public static void main (String [] args){} 

#    print(f"Se ha indexado la información de {len(datos)} países con éxito.")
    #BUSCADOR
#
#   print("-"*50)
#    buscar = input("¿Qué país quieres buscar?: ").strip().capitalize()
#    if buscar in datos:
#        print(f"\nINFORMACIÓN DE: {buscar}")
#        print(datos[buscar])
#    else:
#        print(f"No se ha encontrado información para {buscar}")
 #   
#else:
 #   print(f"No se ha encotrado el archivo en:{rutaReal}") """



