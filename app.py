import pandas as pd
import streamlit as st

from erp_parser import (SIN_DATO, calcular_kpis, generar_csv, generar_excel,
                        procesar_pdf)

st.set_page_config(page_title="Papelera Entre Ríos – PDF a Excel", page_icon="📄", layout="wide")

st.title("📄 Papelera Entre Ríos S.A.")
st.subheader("Convertidor de reportes de stock del ERP (PDF → Excel / CSV)")
st.write("Subí el PDF de **Stock Puro** o de **Stock Destinado a Clientes**. "
         "El tipo de reporte se detecta automáticamente.")

archivos = st.file_uploader("Subir PDF del ERP", type=["pdf"], accept_multiple_files=True)

if "resultados" not in st.session_state:
    st.session_state.resultados = {}

if archivos and st.button("Procesar", type="primary"):
    st.session_state.resultados = {}
    for f in archivos:
        try:
            with st.spinner(f"Procesando {f.name}..."):
                st.session_state.resultados[f.name] = procesar_pdf(f)
        except Exception as e:  # noqa: BLE001
            st.session_state.resultados[f.name] = {"tipo": "error", "mensaje": str(e)}


def fmt(n, dec=0):
    return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False)
def _excel(df, tipo, fecha):
    return generar_excel(df, tipo, fecha)


def filtro_multi(df, col, etiqueta, clave):
    """Multiselect con todo seleccionado por defecto. Los blancos se ven como (sin dato)."""
    serie = df[col]
    if serie.dtype == object:
        serie = serie.replace("", SIN_DATO)
    opciones = sorted(serie.dropna().unique().tolist())
    sel = st.multiselect(etiqueta, opciones, default=opciones, key=clave)
    return serie.isin(sel)


for nombre, r in st.session_state.resultados.items():
    st.divider()
    st.markdown(f"### {nombre}")
    if r["tipo"] == "error":
        st.error(r["mensaje"])
        continue

    tipo, df, fc = r["tipo"], r["df"], r["fecha_corte"]
    etiqueta = "Stock Puro" if tipo == "puro" else "Stock Destinado a Clientes"
    st.caption(f"Reporte detectado: **{etiqueta}** · Fecha de corte: "
               f"{fc.strftime('%d/%m/%Y') if fc else '—'} · {len(df)} filas leídas")

    if r["avisos"]:
        st.warning("**Revisar:**\n\n- " + "\n- ".join(r["avisos"]))
    else:
        st.success(f"Control OK: {fmt(df['Unidades'].sum())} unidades y {fmt(df['Kilos'].sum())} kg "
                   "coinciden con los totales del reporte.")

    # ---------------- Filtros ----------------
    with st.expander("🔎 Filtros", expanded=True):
        c1, c2, c3 = st.columns(3)
        mask = pd.Series(True, index=df.index)
        with c1:
            mask &= filtro_multi(df, "Calidad", "Calidad", f"cal_{nombre}")
            mask &= filtro_multi(df, "Producto", "Producto", f"prod_{nombre}")
        with c2:
            mask &= filtro_multi(df, "Gramaje", "Gramaje", f"gram_{nombre}")
            mask &= filtro_multi(df, "Alistamiento", "Alistamiento", f"ali_{nombre}")
        with c3:
            cmin, cmax = float(df["Corte"].min()), float(df["Corte"].max())
            if cmin < cmax:
                rango = st.slider("Corte (cm)", cmin, cmax, (cmin, cmax), step=0.5, key=f"corte_{nombre}")
                mask &= df["Corte"].between(*rango)
            else:
                st.caption(f"Corte único: {cmin:g} cm")
            if tipo == "destinado":
                mask &= filtro_multi(df, "Cliente", "Cliente", f"cli_{nombre}")

    fdf = df[mask].reset_index(drop=True)
    if fdf.empty:
        st.info("Ningún registro cumple los filtros elegidos.")
        continue
    if len(fdf) < len(df):
        st.caption(f"Mostrando {len(fdf)} de {len(df)} filas. Las descargas incluyen solo lo filtrado.")

    tab_kpi, tab_datos = st.tabs(["📊 Resumen KPIs", "📋 Datos"])

    with tab_kpi:
        k = calcular_kpis(fdf, fc)
        claves = list(k.keys())
        for i in range(0, len(claves), 4):
            cols = st.columns(4)
            for col, clave in zip(cols, claves[i:i + 4]):
                v = k[clave]
                if clave.startswith("%"):
                    txt = f"{v * 100:.1f}%".replace(".", ",")
                elif "Valorización" in clave:
                    txt = fmt(v, 2)
                elif isinstance(v, float):
                    txt = fmt(v, 1)
                else:
                    txt = fmt(v)
                col.metric(clave, txt)
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**Kilos por producto**")
            st.bar_chart(fdf.assign(Producto=fdf["Producto"].replace("", SIN_DATO))
                         .groupby("Producto")["Kilos"].sum())
        with g2:
            st.markdown("**Kilos por calidad**")
            st.bar_chart(fdf.assign(Calidad=fdf["Calidad"].replace("", SIN_DATO))
                         .groupby("Calidad")["Kilos"].sum())

    with tab_datos:
        st.dataframe(fdf, hide_index=True,
                     column_config={"Fecha": st.column_config.DateColumn(format="DD/MM/YYYY")})

    base = nombre.rsplit(".", 1)[0]
    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button("⬇️ Descargar Excel", data=_excel(fdf, tipo, fc),
                       file_name=f"{base}.xlsx", key=f"xl_{nombre}",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇️ Descargar CSV", data=generar_csv(fdf),
                       file_name=f"{base}.csv", mime="text/csv", key=f"csv_{nombre}")
