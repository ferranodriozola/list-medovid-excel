import streamlit as st
import pandas as pd
import requests
import certifi
from io import BytesIO
from xml.sax.saxutils import escape

st.set_page_config(page_title="Llista a xml", page_icon="🛠️")

# Inicialitzar session state per als blocs copiats
if "copied_bloc" not in st.session_state:
    st.session_state.copied_bloc = None

st.markdown('<div id="top" style="position: relative; top: -3.5rem;"></div>', unsafe_allow_html=True)
st.title("XLSX a XML")
st.markdown(
    """
    <style>
    .scroll-top-link {
        position: fixed;
        right: 1.25rem;
        bottom: 2 rem;
        z-index: 9999;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 3rem;
        height: 3rem;
        border-radius: 999px;
        background: #0f766e;
        color: white !important;
        text-decoration: none;
        font-size: 1.4rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
        transition: transform 0.2s ease, box-shadow 0.2s ease, background 0.2s ease;
    }

    .scroll-top-link:hover {
        transform: translateY(-2px);
        background: #115e59;
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.25);
    }
    </style>
    <a class="scroll-top-link" href="#top" title="Puja cap a dalt">↑</a>
    """,
    unsafe_allow_html=True,
)

URL_XLSX = f"https://docs.google.com/spreadsheets/d/{st.secrets['SHEET_ID']}/export?format=xlsx"

PERSON_COLS = {
    'name': 0,           
    'id': 1,
    'woman': 2,          
    'ref_viaf': 3,       
    'ref_2': 4,
    'ref_3': 5,
    'role': 6,           
    'occupation': 7,     
    'birth': 8,          
    'certainty1': 9,     
    'death': 10,          
    'certainty2': 11,     
    'lang_knowledge': 12, 
    'lang_abrev': 13,
    'faith': 14,         
}

PLACE_COLS = {
    'name': 0,      
    'id': 1,        
    'country': 2,   
    'ref_maps': 3,  
    'latitude': 4,  
    'longitude': 5, 
}

# Noms dels blocs de personatges
NOMS_BLOCS = [
    "Mitològics pendents",
    "Puteolanus",
    "Comitibus",
    "Calphurnius",
    "Aleriensis",
    "Accursius",
    "Regius",
    "Alii"
]


@st.cache_data
def descarregar_excel(url_xlsx: str) -> bytes:
    resposta = requests.get(url_xlsx, timeout=30, verify=certifi.where())
    resposta.raise_for_status()
    return resposta.content

@st.cache_data
def obtenir_fulls(url_xlsx: str):
    xls = pd.ExcelFile(BytesIO(descarregar_excel(url_xlsx)))
    return xls.sheet_names

@st.cache_data
def llegir_full(url_xlsx: str, full: str):
    return pd.read_excel(BytesIO(descarregar_excel(url_xlsx)), sheet_name=full)

def _text_segura(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def _llista_camp(valor) -> list[str]:
    text = _text_segura(valor)
    if not text:
        return []
    return [item.strip() for item in text.split(",") if item.strip()]

def _etiqueta_opcional(tag: str, contingut: str, cert: str = "") -> str:
    atribut_cert = f' cert="{escape(cert)}"' if cert else ""
    return f'   <{tag}{atribut_cert}>{escape(contingut)}</{tag}>'


def agrupar_personatges_per_blocs(df_filtrat: pd.DataFrame) -> list[list[pd.Series]]:
    """Agrupa els personatges en blocs separats per files buides."""
    blocs = []
    bloc_actual = []
    
    for idx, fila in df_filtrat.iterrows():
        id_valor = _text_segura(fila.iloc[PERSON_COLS['id']])
        name_valor = _text_segura(fila.iloc[PERSON_COLS['name']])
        
        # Si és una fila vàlida, afegeix al bloc actual
        if id_valor and name_valor:
            bloc_actual.append(fila)
        # Si és una fila buida i hi ha un bloc en construcció, guarda'l
        elif bloc_actual:
            blocs.append(bloc_actual)
            bloc_actual = []
    
    # Afegeix el darrer bloc si existeix
    if bloc_actual:
        blocs.append(bloc_actual)
    
    return blocs

def renderitzar_font_dades(url_xlsx: str, prefix_clau: str) -> None:
    if url_xlsx == URL_XLSX:
        fulls_disponibles = obtenir_fulls(url_xlsx)[:2]

        if not fulls_disponibles:
            st.error("No s'han trobat fulls disponibles a l'Excel.")
            return

        full_seleccionat = st.selectbox(
            "Selecciona el full de l'Excel:",
            fulls_disponibles,
            key=f"{prefix_clau}_full",
        )

        if st.button('Forçar recàrrega', key=f"{prefix_clau}_reload"):
            descarregar_excel.clear()
            obtenir_fulls.clear()
            llegir_full.clear()
            if hasattr(st, 'rerun'):
                st.rerun()
            else:
                st.stop()

        with st.spinner('Processant dades...'):
            try:
                df = llegir_full(url_xlsx, full_seleccionat)

                df_filtrat = df.iloc[0:1000, 0:15]

                if full_seleccionat == 'listPerson':
                    st.subheader("Personatges en format XML")

                    blocs = agrupar_personatges_per_blocs(df_filtrat)

                    if not blocs:
                        st.warning("No hi ha personatges vàlids al full seleccionat.")
                    else:
                        for num_bloc, bloc in enumerate(blocs, 1):
                            # Usar el nom de la llista NOMS_BLOCS o el número per defecte
                            nom_bloc = NOMS_BLOCS[num_bloc - 1] if num_bloc - 1 < len(NOMS_BLOCS) else f"Bloc {num_bloc}"
                            
                            with st.expander(f"{nom_bloc} ({len(bloc)} personatges)", expanded=True):
                                # Botó per copiar tot el bloc
                                xmls_bloc = [construir_person_xml(fila) for fila in bloc]
                                xml_complet_bloc = "\n".join(xmls_bloc)
                                
                                col1, col2, col3 = st.columns([1.5, 2, 2])
                                with col1:
                                    if st.button(
                                        "Copiar bloc",
                                        key=f"{prefix_clau}_copy_bloc_{num_bloc}",
                                        help="Copia tots els XML del bloc"
                                    ):
                                        st.session_state.copied_bloc = num_bloc
                                        # Copiar al portapapeles via JavaScript
                                        st.write(f"**{nom_bloc} copiat!** ({len(bloc)} personatges)")
                                
                                with col2:
                                    st.metric("Personatges", len(bloc))
                                
                                # Mostrar l'XML complet en un bloc de codi
                                st.code(xml_complet_bloc, language='xml')
                                
                                # Mostrar cada personatge individualment
                                with st.expander("Veure personatges individuals"):
                                    for idx, fila in enumerate(bloc, 1):
                                        nom = _text_segura(fila.iloc[PERSON_COLS['name']]) or '(sense nom)'
                                        xml_id = _text_segura(fila.iloc[PERSON_COLS['id']]) or '(sense id)'
                                        st.markdown(f"**{idx}. {nom} ({xml_id})**")
                                        st.code(construir_person_xml(fila), language='xml')

                elif full_seleccionat == 'listPlace':
                    st.subheader("Llocs en format XML")

                    files_valides = df_filtrat.copy()
                    files_valides = files_valides[files_valides.iloc[:, PLACE_COLS['id']].notna()]
                    files_valides = files_valides[files_valides.iloc[:, PLACE_COLS['name']].notna()]

                    if files_valides.empty:
                        st.warning("No hi ha llocs vàlids al full seleccionat.")
                    else:
                        for _, fila in files_valides.iterrows():
                            nom = _text_segura(fila.iloc[PLACE_COLS['name']]) or '(sense nom)'
                            xml_id = _text_segura(fila.iloc[PLACE_COLS['id']]) or nom
                            st.markdown(f"**{nom} ({xml_id})**")
                            st.code(construir_place_xml(fila), language='xml')

                else:
                    st.warning("Full no reconegut. Prova amb listPerson o listPlace.")

            except Exception as e:
                st.error(f"S'ha produït un error en la connexió: {e}")

def construir_person_xml(fila: pd.Series) -> str:
    nom = _text_segura(fila.iloc[PERSON_COLS['name']])
    xml_id = _text_segura(fila.iloc[PERSON_COLS['id']])
    woman = _text_segura(fila.iloc[PERSON_COLS['woman']])
    role = _text_segura(fila.iloc[PERSON_COLS['role']])
    ref = _text_segura(fila.iloc[PERSON_COLS['ref_viaf']])
    ref_2 = _text_segura(fila.iloc[PERSON_COLS['ref_2']])
    ref_3 = _text_segura(fila.iloc[PERSON_COLS['ref_3']])

    ocupacions = _llista_camp(fila.iloc[PERSON_COLS['occupation']])
    birth = _text_segura(fila.iloc[PERSON_COLS['birth']])
    death = _text_segura(fila.iloc[PERSON_COLS['death']])
    cert_birth = _text_segura(fila.iloc[PERSON_COLS['certainty1']])
    cert_death = _text_segura(fila.iloc[PERSON_COLS['certainty2']])
    lang_knowledge = _llista_camp(fila.iloc[PERSON_COLS['lang_knowledge']])
    lang_abrev = _llista_camp(fila.iloc[PERSON_COLS['lang_abrev']])
    faith = _text_segura(fila.iloc[PERSON_COLS['faith']])

    linies = [f'<!-- {escape(nom)} -->\n<person xml:id="{escape(xml_id)}">']

    attrs_persname = []
    if role:
        attrs_persname.append(f'role="{escape(role)}"')
    if woman.lower() == 'woman':
        attrs_persname.append('type="woman"')

    refs = [escape(r) for r in (ref, ref_2, ref_3) if r]
    if refs:
        attrs_persname.append(f'ref="{" ".join(refs)}"')

    attrs_text = f" {' '.join(attrs_persname)}" if attrs_persname else ""
    linies.append(f'   <persName{attrs_text}>{escape(nom)}</persName>')

    for ocupacio in ocupacions:
        linies.append(f'   <occupation>{escape(ocupacio)}</occupation>')

    if birth:
        linies.append(_etiqueta_opcional('birth', birth, cert_birth))
    if death:
        linies.append(_etiqueta_opcional('death', death, cert_death))
    if lang_knowledge:
        linies.append('   <langKnowledge>')
        for idioma, abreviatura in zip(lang_knowledge, lang_abrev):
            linies.append(f'      <langKnown tag="{escape(abreviatura)}">{escape(idioma)}</langKnown>')
        linies.append('   </langKnowledge>')
    if faith:
        linies.append(f'   <faith>{escape(faith)}</faith>')

    linies.append('</person>\n')
    return '\n'.join(linies)

def construir_place_xml(fila: pd.Series) -> str:
    nom = _text_segura(fila.iloc[PLACE_COLS['name']])
    xml_id = _text_segura(fila.iloc[PLACE_COLS['id']]) or nom
    ref = _text_segura(fila.iloc[PLACE_COLS['ref_maps']])
    pais = _text_segura(fila.iloc[PLACE_COLS['country']])
    latitud = _text_segura(fila.iloc[PLACE_COLS['latitude']])
    longitud = _text_segura(fila.iloc[PLACE_COLS['longitude']])

    linies = [f'<!-- {escape(nom)} -->\n<place xml:id="{escape(xml_id)}">']

    attrs_placename = []
    if ref:
        attrs_placename.append(f'ref="{escape(ref)}"')

    attrs_text = f" {' '.join(attrs_placename)}" if attrs_placename else ""
    linies.append(f'   <placeName{attrs_text}>{escape(nom)}</placeName>')

    if pais:
        linies.append(f'   <country>{escape(pais)}</country>')

    if latitud or longitud:
        linies.append('   <location>')
        geo_text = f"{latitud}, {longitud}" if (latitud and longitud) else (latitud or longitud)
        linies.append(f'      <geo>{escape(geo_text)}</geo>')
        linies.append('   </location>')

    linies.append('</place>\n')
    return '\n'.join(linies)

def construir_person_xml_simple(fila: pd.Series) -> str:
    nom = _text_segura(fila.iloc[0])  # Name (columna 0)
    xml_id = _text_segura(fila.iloc[1])  # ID (columna 1)
    
    linies = [f'<!-- {escape(nom)} -->', f'<person xml:id="{escape(xml_id)}">', f'   <persName>{escape(nom)}</persName>', '</person>']
    
    return '\n'.join(linies)

tabs = st.tabs(["SHEET_ID"])

with tabs[0]:
    renderitzar_font_dades(URL_XLSX, "sheet_1")