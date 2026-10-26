from datetime import date, timedelta
import io
import json
import os
import re
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

CONFIG_FILE = "config.json"


# ---------------------------------------------------------
# CLASE PARA NUMERACIÓN DINÁMICA DE PÁGINAS ("Página X de Y")
# ---------------------------------------------------------
class NumberedCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#595959"))
        text_num = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(612 - 20, 15, text_num)


# ---------------------------------------------------------
# FUNCIONES DE PERSISTENCIA Y REINICIO DIARIO
# ---------------------------------------------------------
def obtener_ultimo_correlativo():
    fecha_hoy = date.today().strftime("%Y-%m-%d")

    if not os.path.exists(CONFIG_FILE):
        return 1

    try:
        with open(CONFIG_FILE, "r") as f:
            data = json.load(f)
            ultima_fecha = data.get("ultima_fecha", "")

            if ultima_fecha != fecha_hoy:
                return 1

            return data.get("ultimo_correlativo", 1)
    except Exception:
        return 1


def guardar_ultimo_correlativo(nuevo_correlativo):
    fecha_hoy = date.today().strftime("%Y-%m-%d")
    with open(CONFIG_FILE, "w") as f:
        json.dump(
            {
                "ultimo_correlativo": nuevo_correlativo,
                "ultima_fecha": fecha_hoy,
            },
            f,
        )


def cargar_excel_inteligente(file):
    """Detecta la fila donde comienzan las columnas del ERP y limpia espacios en los nombres de columnas."""
    df_temp = pd.read_excel(file, header=None, dtype=str)

    header_row_idx = None
    for idx, row in df_temp.iterrows():
        row_values = [str(val).strip().upper() for val in row.dropna().values]
        if "KOPRCT" in row_values or "FMPR" in row_values:
            header_row_idx = idx
            break

    if header_row_idx is not None:
        file.seek(0)
        df = pd.read_excel(file, skiprows=header_row_idx, dtype=str)
    else:
        file.seek(0)
        df = pd.read_excel(file, dtype=str)

    # Limpiar espacios en blanco al inicio y final de los nombres de columnas
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def calcular_desglose_cajas(cap1, cap2, descripcion=""):
    """Calcula con precisión exacta Cajas Enteras y Unidades Sueltas."""
    try:
        cap1 = float(cap1)
        if cap1 <= 0:
            return 0, 0

        cap2 = float(cap2) if cap2 is not None else 0.0
        unidades_por_caja = 0

        # 1. Intentar deducir desde CAPRCO2 (ej: 27 / 2.7 = 10)
        if cap2 > 0:
            unidades_por_caja = round(cap1 / round(cap2, 4))

        # 2. Respaldo: Si CAPRCO2 era 0 o 1, buscar factor en la descripción (ej: "10X1 KG" -> 10)
        if unidades_por_caja <= 1 and descripcion:
            match = re.search(r'(\d+)\s*[xX]', str(descripcion))
            if match:
                unidades_por_caja = int(match.group(1))

        # 3. Calcular desglose
        if unidades_por_caja > 1:
            cajas_enteras = int(cap1 // unidades_por_caja)
            unidades_sueltas = int(round(cap1 % unidades_por_caja))
            return cajas_enteras, unidades_sueltas

        return 0, int(round(cap1))
    except Exception:
        return 0, int(round(cap1))


def generar_pdf_picking_por_hojas(list_grupos_data, obs_pde):
    """Genera un PDF con la Comuna/Observación ubicada en un recuadro destacado arriba a la derecha."""
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=30,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "GroupTitle",
        parent=styles["Heading1"],
        fontSize=12,
        leading=14,
        textColor=colors.HexColor("#1F4E78"),
        fontName="Helvetica-Bold",
        spaceAfter=2,
    )

    sub_info_style = ParagraphStyle(
        "SubInfo",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#595959"),
        fontName="Helvetica-Oblique",
    )

    badge_label_style = ParagraphStyle(
        "BadgeLabel",
        fontSize=7,
        leading=8,
        textColor=colors.HexColor("#1F4E78"),
        fontName="Helvetica-Bold",
        alignment=1,
    )

    badge_text_style = ParagraphStyle(
        "BadgeText",
        fontSize=11,
        leading=13,
        textColor=colors.HexColor("#000000"),
        fontName="Helvetica-Bold",
        alignment=1,
    )

    cat_banner_style = ParagraphStyle(
        "CatBanner",
        parent=styles["Heading2"],
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#1F4E78"),
        fontName="Helvetica-Bold",
        spaceBefore=4,
        spaceAfter=2,
    )

    hdr_text_style = ParagraphStyle(
        "HdrText",
        fontSize=7,
        leading=8.5,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )
    body_text_style = ParagraphStyle(
        "BodyText", fontSize=7, leading=8.5, fontName="Helvetica"
    )
    body_right_style = ParagraphStyle(
        "BodyRight",
        fontSize=7,
        leading=8.5,
        fontName="Helvetica",
        alignment=2,
    )
    body_bold_right_style = ParagraphStyle(
        "BodyBoldRight",
        fontSize=7,
        leading=8.5,
        fontName="Helvetica-Bold",
        alignment=2,
    )
    total_right_style = ParagraphStyle(
        "TotalRight",
        fontSize=7.5,
        leading=9,
        fontName="Helvetica-Bold",
        alignment=2,
    )

    elements = []
    total_grupos = len(list_grupos_data)

    for idx_g, (id_consol, df_g) in enumerate(list_grupos_data):
        left_flowables = [
            Paragraph(
                f"📦 DETALLE DE PICKING | GRUPO: {id_consol}", title_style
            ),
            Spacer(1, 2),
            Paragraph(
                f"Fecha Emisión: {date.today().strftime('%d/%m/%Y')}",
                sub_info_style,
            ),
        ]

        badge_content = [
            Paragraph("COMUNA / RUTA DE ATENCIÓN", badge_label_style),
            Spacer(1, 2),
            Paragraph(f"{obs_pde.upper()}", badge_text_style),
        ]

        badge_table = Table([[badge_content]], colWidths=[180])
        badge_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F2F5F9")),
                    ("BOX", (0, 0), (-1, -1), 1.2, colors.HexColor("#1F4E78")),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )

        header_table = Table(
            [[left_flowables, badge_table]], colWidths=[382, 190]
        )
        header_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )

        elements.append(header_table)
        elements.append(Spacer(1, 10))

        categorias_unicas = (
            df_g[
                ["Abreviación Categoría", "Nombre Categoría"]
            ].drop_duplicates()
            .sort_values(by="Abreviación Categoría")
        )

        for _, cat in categorias_unicas.iterrows():
            f_code = cat["Abreviación Categoría"]
            f_name = cat["Nombre Categoría"]

            bloque_categoria = []
            bloque_categoria.append(
                Paragraph(
                    f"🏷️ CATEGORÍA: {f_code} - {f_name}", cat_banner_style
                )
            )

            table_data = [
                [
                    Paragraph("Abr.", hdr_text_style),
                    Paragraph("Categoría", hdr_text_style),
                    Paragraph("SKU", hdr_text_style),
                    Paragraph("Descripción Producto", hdr_text_style),
                    Paragraph("Cant. Total", hdr_text_style),
                    Paragraph("Cant. Sec.", hdr_text_style),
                    Paragraph("Cajas", hdr_text_style),
                    Paragraph("Unid. Sueltas", hdr_text_style),
                ]
            ]

            df_cat_prods = df_g[df_g["Abreviación Categoría"] == f_code].copy()

            df_cat_prods["SKU_NUM"] = pd.to_numeric(
                df_cat_prods["Código Producto"], errors="coerce"
            )
            df_cat_prods = df_cat_prods.sort_values(
                by=["SKU_NUM", "Descripción Producto"], ascending=[True, True]
            ).drop(columns=["SKU_NUM"])

            sum_cap1 = 0
            sum_cap2 = 0
            sum_cajas = 0
            sum_sueltas = 0

            for _, r in df_cat_prods.iterrows():
                v1 = int(r["Cantidad Primaria (CAPRCO1)"])
                v2 = float(r["Cantidad Secundaria (CAPRCO2)"])
                desc = str(r["Descripción Producto"])

                cajas_ent, unid_sueltas = calcular_desglose_cajas(v1, v2, desc)

                sum_cap1 += v1
                sum_cap2 += v2
                sum_cajas += cajas_ent
                sum_sueltas += unid_sueltas

                table_data.append(
                    [
                        Paragraph(
                            str(r["Abreviación Categoría"]), body_text_style
                        ),
                        Paragraph(str(r["Nombre Categoría"]), body_text_style),
                        Paragraph(str(r["Código Producto"]), body_text_style),
                        Paragraph(
                            str(r["Descripción Producto"]), body_text_style
                        ),
                        Paragraph(f"{v1:,}", body_right_style),
                        Paragraph(f"{v2:,.2f}", body_right_style),
                        Paragraph(f"{cajas_ent:,}", body_bold_right_style),
                        Paragraph(f"{unid_sueltas:,}", body_bold_right_style),
                    ]
                )

            table_data.append(
                [
                    Paragraph("", body_text_style),
                    Paragraph("", body_text_style),
                    Paragraph("", body_text_style),
                    Paragraph(f"TOTAL {f_name}:", total_right_style),
                    Paragraph(f"{sum_cap1:,}", total_right_style),
                    Paragraph(f"{sum_cap2:,.2f}", total_right_style),
                    Paragraph(f"{sum_cajas:,}", total_right_style),
                    Paragraph(f"{sum_sueltas:,}", total_right_style),
                ]
            )

            t = Table(
                table_data,
                colWidths=[35, 70, 50, 197, 55, 55, 55, 55],
                repeatRows=1,
            )
            t.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5597")),
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ("TOPPADDING", (0, 0), (-1, -1), 2),
                        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F2F2F2")),
                    ]
                )
            )

            bloque_categoria.append(t)
            bloque_categoria.append(Spacer(1, 6))

            elements.append(KeepTogether(bloque_categoria))

        if idx_g < total_grupos - 1:
            elements.append(PageBreak())

    doc.build(elements, canvasmaker=NumberedCanvas)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


# ---------------------------------------------------------
# CONEXIÓN CON FIREBASE FIRESTORE
# ---------------------------------------------------------
@st.cache_resource
def get_firestore_client():
    if not firebase_admin._apps:
        cred = credentials.Certificate("firebase_credentials.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()


def subir_consolidado_a_firebase(df_final):
    """Carga los registros en Firestore en lotes incluyendo superfamilia, desglose de cajas y descripción para la tablet."""
    db = get_firestore_client()
    batch = db.batch()
    operaciones = 0
    total_registros = len(df_final)

    for _, row in df_final.iterrows():
        doc_id = f"{row['PDE_num_doc']}_{row['PDE_lin_doc']}"
        doc_ref = db.collection("despachos").document(doc_id)

        cant_teorica = int(row["PDE_pdt_mat"])

        data = {
            "cod_tdo": str(row["PDE_cod_tdo"]),
            "num_doc": str(row["PDE_num_doc"]),
            "lin_doc": int(row["PDE_lin_doc"]),
            "fec_emi": str(row["PDE_fec_emi"]),
            "fec_ent": str(row["PDE_fec_ent"]),
            "cod_cli": str(row["PDE_cod_cli"]),
            "des_cli": str(row["PDE_des_cli"]),
            "cod_mat": str(row["PDE_cod_mat"]),
            # Metadatos para orden y agrupamiento en la Tablet
            "des_mat": str(row.get("PDE_nom_mat", "")),
            "cod_fam": str(row.get("PDE_cod_fam", "")),
            "des_fam": str(row.get("PDE_nom_fam", "")),
            "superfam": str(row.get("PDE_superfam", "GENERAL")).upper(),
            "pdt_mat": cant_teorica,
            # Desglose físico de Cajas y Unidades Sueltas (igual que en PDF)
            "cajas": int(row.get("PDE_cajas", 0)),
            "sueltas": int(row.get("PDE_sueltas", cant_teorica)),
            "obs_pde": str(row["PDE_obs_pde"]),
            "zon_cli": str(row["PDE_zon_cli"]),
            "ingreso_camion": False,
            "estado_carga": "PENDIENTE",
            "cant_real": cant_teorica,
            "obs_tablet": "",
            "cerrado": False,
            "fecha_cierre": None,
            "cargado_en": firestore.SERVER_TIMESTAMP,
        }

        batch.set(doc_ref, data, merge=True)
        operaciones += 1

        if operaciones == 400:
            batch.commit()
            batch = db.batch()
            operaciones = 0

    if operaciones > 0:
        batch.commit()

    return total_registros


def obtener_consolidados_registrados():
    """Consulta Firestore y agrupa consolidados por Ruta/Comuna y Fecha."""
    db = get_firestore_client()
    docs = db.collection("despachos").stream()
    registros = [d.to_dict() for d in docs]

    if not registros:
        return pd.DataFrame()

    df = pd.DataFrame(registros)
    if "obs_pde" not in df.columns or "fec_emi" not in df.columns:
        return pd.DataFrame()

    if "cerrado" not in df.columns:
        df["cerrado"] = False

    resumen = (
        df.groupby(["obs_pde", "fec_emi", "num_doc"], as_index=False)
        .agg(
            total_items=("cod_mat", "count"),
            cerrado=("cerrado", lambda x: (x == True).all()),
            revisados=(
                "estado_carga",
                lambda x: (x.isin(["OK", "AJUSTE", "NO_SALIO"])).sum(),
            ),
        )
    )
    return resumen


def reabrir_consolidado_en_firebase(obs_pde_sel, fec_emi_sel, num_doc_sel=None):
    """Reabre todos los documentos de la ruta en Firestore quitando la marca de cerrado."""
    db = get_firestore_client()
    query = (
        db.collection("despachos")
        .where("obs_pde", "==", obs_pde_sel)
        .where("fec_emi", "==", fec_emi_sel)
    )

    docs = query.stream()
    batch = db.batch()
    conteo = 0

    for doc_snap in docs:
        batch.update(doc_snap.reference, {"cerrado": False, "fecha_cierre": None})
        conteo += 1
        if conteo % 400 == 0:
            batch.commit()
            batch = db.batch()

    if conteo % 400 != 0:
        batch.commit()

    return conteo


# ---------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Generador de Consolidados Picking",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    div[data-testid="stTextInput"]:has(input[aria-label*="Comuna / Observación"]) input:placeholder-shown {
        border: 2px solid #FF4B4B !important;
        background-color: #2D1A1A !important;
        box-shadow: 0 0 8px rgba(255, 75, 75, 0.6) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 Generador de Consolidados e Integración Picking")

tab_generar, tab_administrar = st.tabs(["🚀 Generar Consolidados", "🔓 Administrar / Reabrir Consolidados"])

# ==============================================================================
# TAB 1: GENERAR CONSOLIDADOS (PICKING / DESPACHO)
# ==============================================================================
with tab_generar:
    with st.container():
        st.markdown("### ⚙️ Configuración y Carga de Archivo")

        col_top1, col_top2 = st.columns([1.5, 2.5])

        with col_top1:
            archivo_excel = st.file_uploader(
                "📁 Cargar Excel de Entrada", type=["xlsx", "xls"]
            )

        with col_top2:
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                cod_tdo = st.text_input("Código Tipo Doc", value="NVI")
                zon_cli = st.text_input("Zona Cliente", value="zona")

            with col_p2:
                fec_emi = st.date_input("Fecha Emisión", value=date.today())
                fec_ent = st.date_input(
                    "Fecha Entrega", value=date.today() + timedelta(days=1)
                )

            with col_p3:
                obs_pde = st.text_input(
                    "Comuna / Observación *",
                    value="",
                    placeholder="⚠️ EJ: SAN CARLOS 11",
                )

                correlativo_actual_base = obtener_ultimo_correlativo()
                nuevo_corr_manual = st.number_input(
                    "Correlativo HOY:",
                    min_value=1,
                    value=correlativo_actual_base,
                    step=1,
                )

                if nuevo_corr_manual != correlativo_actual_base:
                    guardar_ultimo_correlativo(nuevo_corr_manual)

    st.divider()

    # ---------------------------------------------------------
    # FLUJO PRINCIPAL DE PROCESAMIENTO
    # ---------------------------------------------------------
    if archivo_excel is not None:
        try:
            df_original = cargar_excel_inteligente(archivo_excel)

            df_original["CAPRCO1"] = pd.to_numeric(
                df_original["CAPRCO1"], errors="coerce"
            ).fillna(0)
            df_original["CAPRCO2"] = pd.to_numeric(
                df_original["CAPRCO2"], errors="coerce"
            ).fillna(0)

            with st.expander("👀 Ver datos de entrada cargados", expanded=False):
                st.dataframe(df_original.head(10), use_container_width=True)

            familias_df = df_original[["FMPR", "NOFMPR"]].drop_duplicates()
            familias_disponibles = familias_df.apply(
                lambda x: f"{x['FMPR']} - {x['NOFMPR']}", axis=1
            ).tolist()

            st.subheader("🧩 1. Definición de Grupos / Consolidados")

            num_grupos = st.number_input(
                "¿Cuántos grupos o consolidados deseas crear?",
                min_value=1,
                max_value=10,
                value=3,
                step=1,
            )

            familias_seleccionadas_por_grupo = {}
            todas_las_seleccionadas = set()

            for i in range(int(num_grupos)):
                key = f"fam_g_{i}"
                sel = st.session_state.get(key, [])
                familias_seleccionadas_por_grupo[i] = sel
                todas_las_seleccionadas.update(sel)

            familias_faltantes = [
                f for f in familias_disponibles if f not in todas_las_seleccionadas
            ]

            observacion_valida = bool(obs_pde.strip())

            if familias_faltantes:
                st.warning(
                    f"⚠️ **Faltan {len(familias_faltantes)} categorías por asignar**: "
                    + ", ".join([f.split(" - ")[1] for f in familias_faltantes])
                )

            if not observacion_valida:
                st.error(
                    "🛑 **ATENCIÓN: Debe ingresar el nombre de la Comuna / Observación (marcado en rojo arriba) para continuar.**"
                )

            if not familias_faltantes and observacion_valida:
                st.success(
                    "✅ **¡Todas las categorías han sido asignadas y los parámetros están completos!**"
                )

            list_dfs_procesados = []
            list_grupos_data_export = []
            fecha_id_str = date.today().strftime("%d%m%Y")
            correlativo_inicial = obtener_ultimo_correlativo()

            # Bucle de construcción de cada grupo
            for i in range(int(num_grupos)):
                num_consol = i + 1
                corr_grupo = correlativo_inicial + i
                id_autogenerado = f"{corr_grupo}{fecha_id_str}"

                st.markdown(f"### 📦 Grupo {num_consol}")

                col1, col2 = st.columns([1, 2])
                with col1:
                    nom_consol = st.text_input(
                        f"Descripción Grupo {num_consol}",
                        value=id_autogenerado,
                        key=f"nom_g_{i}",
                    )
                    id_consol = st.text_input(
                        f"ID Documento Grupo {num_consol}",
                        value=id_autogenerado,
                        key=f"id_g_{i}",
                    )

                with col2:
                    mis_seleccionadas = familias_seleccionadas_por_grupo[i]
                    ocupadas_por_otros = todas_las_seleccionadas - set(
                        mis_seleccionadas
                    )

                    opciones_disponibles = [
                        f
                        for f in familias_disponibles
                        if f not in ocupadas_por_otros
                    ]

                    familias_sel = st.multiselect(
                        f"Selecciona las familias de productos para {id_consol}:",
                        options=opciones_disponibles,
                        key=f"fam_g_{i}",
                    )

                # Procesamiento de CADA grupo dentro del bucle
                if familias_sel:
                    cods_fam = [f.split(" - ")[0] for f in familias_sel]

                    df_grupo = df_original[
                        df_original["FMPR"].isin(cods_fam)
                    ].copy()

                    # Asegurar detección exacta y flexible de SUPERFAM
                    col_super_encontrada = next(
                        (
                            c
                            for c in df_grupo.columns
                            if "SUPERFAM" in c or "FMSU" in c
                        ),
                        None,
                    )

                    if col_super_encontrada:
                        df_grupo["SUPERFAM"] = (
                            df_grupo[col_super_encontrada]
                            .fillna("GENERAL")
                            .astype(str)
                            .str.strip()
                            .str.upper()
                        )
                    else:
                        df_grupo["SUPERFAM"] = (
                            df_grupo["FMPR"]
                            .fillna("GENERAL")
                            .astype(str)
                            .str.strip()
                            .str.upper()
                        )

                    # Reemplazar valores vacíos
                    df_grupo["SUPERFAM"] = df_grupo["SUPERFAM"].replace(
                        ["", "NAN", "NONE"], "GENERAL"
                    )

                    # Agrupar consolidando CAPRCO1 y CAPRCO2 preservando SUPERFAM
                    df_grupo = (
                        df_grupo.groupby(
                            ["KOPRCT", "NOKOPRCT", "FMPR", "NOFMPR", "SUPERFAM"],
                            as_index=False,
                            sort=False,
                        )[["CAPRCO1", "CAPRCO2"]]
                        .sum()
                        .reset_index(drop=True)
                    )

                    # Ordenar 1° SUPERFAM (A-Z) y 2° SKU de menor a mayor
                    df_grupo["KOPRCT_NUM"] = pd.to_numeric(
                        df_grupo["KOPRCT"].str.extract(r"(\d+)", expand=False),
                        errors="coerce",
                    ).fillna(0)
                    df_grupo = df_grupo.sort_values(
                        by=["SUPERFAM", "KOPRCT_NUM", "KOPRCT"],
                        ascending=[True, True, True],
                    ).drop(columns=["KOPRCT_NUM"])
                    df_grupo = df_grupo.reset_index(drop=True)

                    # CSV Oficial ERP + columnas para Firebase
                    df_csv_grupo = df_grupo.copy()
                    df_csv_grupo["PDE_cod_tdo"] = cod_tdo
                    df_csv_grupo["PDE_num_doc"] = id_consol
                    df_csv_grupo["PDE_lin_doc"] = df_csv_grupo.index + 1
                    df_csv_grupo["PDE_fec_emi"] = fec_emi.strftime("%Y-%m-%d")
                    df_csv_grupo["PDE_fec_ent"] = fec_ent.strftime("%Y-%m-%d")
                    df_csv_grupo["PDE_cod_cli"] = 1
                    df_csv_grupo["PDE_des_cli"] = nom_consol
                    df_csv_grupo["PDE_cod_mat"] = df_csv_grupo["KOPRCT"]
                    df_csv_grupo["PDE_pdt_mat"] = df_csv_grupo["CAPRCO1"].astype(
                        int
                    )
                    df_csv_grupo["PDE_obs_pde"] = obs_pde
                    df_csv_grupo["PDE_zon_cli"] = zon_cli
                    df_csv_grupo["PDE_val_pdt"] = 1

                    # Columnas auxiliares para la Tablet
                    df_csv_grupo["PDE_nom_mat"] = df_grupo["NOKOPRCT"]
                    df_csv_grupo["PDE_cod_fam"] = df_grupo["FMPR"]
                    df_csv_grupo["PDE_nom_fam"] = df_grupo["NOFMPR"]
                    df_csv_grupo["PDE_superfam"] = df_grupo["SUPERFAM"]

                    # Calcular desglose exacto de cajas y sueltas
                    desglose = df_grupo.apply(
                        lambda r: calcular_desglose_cajas(r["CAPRCO1"], r["CAPRCO2"], r["NOKOPRCT"]),
                        axis=1,
                    )
                    df_csv_grupo["PDE_cajas"] = [d[0] for d in desglose]
                    df_csv_grupo["PDE_sueltas"] = [d[1] for d in desglose]

                    columnas_csv_totales = [
                        "PDE_cod_tdo",
                        "PDE_num_doc",
                        "PDE_lin_doc",
                        "PDE_fec_emi",
                        "PDE_fec_ent",
                        "PDE_cod_cli",
                        "PDE_des_cli",
                        "PDE_cod_mat",
                        "PDE_pdt_mat",
                        "PDE_obs_pde",
                        "PDE_zon_cli",
                        "PDE_val_pdt",
                        "PDE_nom_mat",
                        "PDE_cod_fam",
                        "PDE_nom_fam",
                        "PDE_superfam",
                        "PDE_cajas",
                        "PDE_sueltas",
                    ]

                    df_resultado_grupo = df_csv_grupo[columnas_csv_totales]
                    list_dfs_procesados.append(df_resultado_grupo)

                    # Reporte PDF
                    df_export_g = df_grupo[
                        [
                            "FMPR",
                            "NOFMPR",
                            "KOPRCT",
                            "NOKOPRCT",
                            "CAPRCO1",
                            "CAPRCO2",
                        ]
                    ].copy()
                    df_export_g = df_export_g.rename(
                        columns={
                            "FMPR": "Abreviación Categoría",
                            "NOFMPR": "Nombre Categoría",
                            "KOPRCT": "Código Producto",
                            "NOKOPRCT": "Descripción Producto",
                            "CAPRCO1": "Cantidad Primaria (CAPRCO1)",
                            "CAPRCO2": "Cantidad Secundaria (CAPRCO2)",
                        }
                    )

                    list_grupos_data_export.append((id_consol, df_export_g))

                    st.caption(
                        f"✓ **{len(df_resultado_grupo)}** ítems agregados a {id_consol}."
                    )

                    with st.expander(
                        f"🔍 Ver listado de {len(df_resultado_grupo)} productos asignados a {id_consol}",
                        expanded=False,
                    ):
                        st.dataframe(df_export_g, use_container_width=True)

                st.divider()

            # ---------------------------------------------------------
            # UNIFICACIÓN Y EXPORTACIÓN (FUERA DEL FOR LOOP)
            # ---------------------------------------------------------
            if list_dfs_procesados:
                df_final = pd.concat(list_dfs_procesados, ignore_index=True)

                columnas_erp = [
                    "PDE_cod_tdo",
                    "PDE_num_doc",
                    "PDE_lin_doc",
                    "PDE_fec_emi",
                    "PDE_fec_ent",
                    "PDE_cod_cli",
                    "PDE_des_cli",
                    "PDE_cod_mat",
                    "PDE_pdt_mat",
                    "PDE_obs_pde",
                    "PDE_zon_cli",
                    "PDE_val_pdt",
                ]
                df_para_csv_erp = df_final[columnas_erp]

                st.subheader("📋 2. Previsualización y Exportación de Resultados")
                st.dataframe(df_para_csv_erp, use_container_width=True)

                csv_data = df_para_csv_erp.to_csv(index=False, sep=";").encode(
                    "utf-8"
                )

                pdf_bytes = (
                    generar_pdf_picking_por_hojas(list_grupos_data_export, obs_pde)
                    if observacion_valida
                    else b""
                )

                def al_descargar():
                    siguiente_correlativo = correlativo_inicial + int(num_grupos)
                    guardar_ultimo_correlativo(siguiente_correlativo)

                puede_descargar = bool(not familias_faltantes and observacion_valida)

                col_btn1, col_btn2, col_btn3 = st.columns(3)

                with col_btn1:
                    st.download_button(
                        label="📥 Descargar CSV para Sistema (ERP)",
                        data=csv_data,
                        file_name=(
                            f"{obs_pde}.csv" if observacion_valida else "salida.csv"
                        ),
                        mime="text/csv",
                        on_click=al_descargar,
                        disabled=not puede_descargar,
                        use_container_width=True,
                    )

                with col_btn2:
                    st.download_button(
                        label="📄 Descargar PDF de Picking",
                        data=pdf_bytes,
                        file_name=(
                            f"Picking_PDF_{obs_pde}.pdf"
                            if observacion_valida
                            else "picking.pdf"
                        ),
                        mime="application/pdf",
                        disabled=not puede_descargar,
                        use_container_width=True,
                    )

                with col_btn3:
                    if st.button(
                        "☁️ Enviar a Tablets (Firebase)",
                        disabled=not puede_descargar,
                        use_container_width=True,
                        type="primary",
                    ):
                        with st.spinner("Sincronizando con Firebase..."):
                            total = subir_consolidado_a_firebase(df_final)
                            st.success(f"✅ ¡{total} registros subidos con éxito!")

        except Exception as e:
            st.error(f"Error al procesar la información: {e}")
    else:
        st.info(
            "👆 Por favor, carga tu archivo Excel en la barra superior para iniciar."
        )

# ==============================================================================
# TAB 2: ADMINISTRAR / REABRIR CONSOLIDADOS CERRADOS
# ==============================================================================
with tab_administrar:
    st.subheader("🔓 Panel de Gestión y Reapertura de Consolidados")
    st.markdown(
        "Utiliza esta sección para reabrir rutas o consolidados que hayan sido"
        " cerrados por error en la tablet, o para consultar el avance de días"
        " anteriores."
    )

    col_ref, _ = st.columns([1, 4])
    with col_ref:
        recargar_btn = st.button(
            "🔄 Actualizar lista desde Firebase", use_container_width=True
        )

    try:
        df_consolidados = obtener_consolidados_registrados()

        if df_consolidados.empty:
            st.info(
                "No se encontraron consolidados registrados en la base de datos."
            )
        else:
            df_consolidados["label"] = df_consolidados.apply(
                lambda r: (
                    f"{'🔒 [CERRADO]' if r['cerrado'] else '🚚 [ABIERTO]'} "
                    f"{r['obs_pde']} | Fecha: {r['fec_emi']} | Doc: "
                    f"{r['num_doc']} ({r['revisados']}/{r['total_items']} revisados)"
                ),
                axis=1,
            )

            st.dataframe(
                df_consolidados[[
                    "obs_pde",
                    "fec_emi",
                    "num_doc",
                    "total_items",
                    "revisados",
                    "cerrado",
                ]].rename(
                    columns={
                        "obs_pde": "Comuna / Ruta",
                        "fec_emi": "Fecha Emisión",
                        "num_doc": "Doc Consolidado",
                        "total_items": "Total Ítems",
                        "revisados": "Chequeados",
                        "cerrado": "Cerrado en Tablet",
                    }
                ),
                use_container_width=True,
            )

            opciones = df_consolidados["label"].tolist()
            seleccion = st.selectbox(
                "Selecciona el consolidado que deseas gestionar:", opciones
            )

            if seleccion:
                fila_sel = df_consolidados[
                    df_consolidados["label"] == seleccion
                ].iloc[0]
                obs_sel = fila_sel["obs_pde"]
                fec_sel = fila_sel["fec_emi"]
                doc_sel = fila_sel["num_doc"]
                esta_cerrado = fila_sel["cerrado"]

                st.markdown("---")
                col_info, col_accion = st.columns([2, 1])

                with col_info:
                    st.write(f"**Ruta:** `{obs_sel}`")
                    st.write(f"**Fecha:** `{fec_sel}`")
                    st.write(f"**Documento:** `{doc_sel}`")
                    st.write(
                        f"**Estado actual:** {'🔴 CERRADO' if esta_cerrado else '🟢 ABIERTO (Visible en pendientes)'}"
                    )

                with col_accion:
                    if esta_cerrado:
                        if st.button(
                            "🔓 Reabrir este Consolidado",
                            type="primary",
                            use_container_width=True,
                        ):
                            with st.spinner(
                                "Reabriendo consolidado en Firebase..."
                            ):
                                cant = reabrir_consolidado_en_firebase(
                                    obs_sel, fec_sel, doc_sel
                                )
                                st.success(
                                    f"✅ ¡Consolidado reabierto! ({cant} ítems ahora editables en tablet)"
                                )
                                st.rerun()
                    else:
                        st.info(
                            "Este consolidado ya se encuentra abierto y editable en la tablet."
                        )

    except Exception as err:
        st.error(f"Error al consultar Firebase: {err}")