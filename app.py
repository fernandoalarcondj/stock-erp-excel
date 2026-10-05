import altair as alt
import pandas as pd
import streamlit as st

from erp_parser import (SIN_DATO, calcular_kpis, generar_csv, generar_excel,
                        kilos_por_corte, procesar_pdf)

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


def corte_txt(x):
    return f"{x:g}"


@st.cache_data(show_spinner=False)
def _excel(df, tipo, fecha, dias):
    return generar_excel(df, tipo, fecha, dias)


def filtro_lista(df, col, etiqueta, clave):
    """Lista desplegable con opción 'Todos'. Los blancos se ven como (sin dato)."""
    serie = df[col]
    if serie.dtype == object:
        serie = serie.replace("", SIN_DATO)
    opciones = ["Todos"] + sorted(serie.dropna().unique().tolist())
    sel = st.selectbox(etiqueta, opciones, key=clave,
                       format_func=lambda x: x if isinstance(x, str) else corte_txt(x))
    if sel == "Todos":
        return pd.Series(True, index=df.index)
    return serie == sel


def barras(datos, x, y, titulo_y="Kilos", orden=None):
    graf = alt.Chart(datos).mark_bar().encode(
        x=alt.X(f"{x}:N", sort=orden, title=x, axis=alt.Axis(labelAngle=-45)),
        y=alt.Y(f"{y}:Q", title=titulo_y),
        tooltip=[x, alt.Tooltip(f"{y}:Q", format=",.0f")],
    ).properties(height=320)
    st.altair_chart(graf, width="stretch")


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

    # ---------------- Filtros (listas desplegables) ----------------
    st.markdown("**🔎 Filtros**")
    mask = pd.Series(True, index=df.index)
    f1, f2, f3 = st.columns(3)
    with f1:
        mask &= filtro_lista(df, "Calidad", "Calidad", f"cal_{nombre}")
    with f2:
        mask &= filtro_lista(df, "Gramaje", "Gramaje (g/m²)", f"gram_{nombre}")
    with f3:
        mask &= filtro_lista(df, "Corte", "Corte (cm)", f"corte_{nombre}")
    f4, f5, f6 = st.columns(3)
    with f4:
        mask &= filtro_lista(df, "Producto", "Producto", f"prod_{nombre}")
    with f5:
        mask &= filtro_lista(df, "Alistamiento", "Alistamiento", f"ali_{nombre}")
    with f6:
        if tipo == "destinado":
            mask &= filtro_lista(df, "Cliente", "Cliente", f"cli_{nombre}")
        dias = st.number_input("Días para considerar stock antiguo", 1, 3650, 90, key=f"dias_{nombre}")

    fdf = df[mask].reset_index(drop=True)
    if fdf.empty:
        st.info("Ningún registro cumple los filtros elegidos.")
        continue
    if len(fdf) < len(df):
        st.caption(f"Mostrando {len(fdf)} de {len(df)} filas. Las descargas incluyen solo lo filtrado.")

    tab_kpi, tab_datos = st.tabs(["📊 Resumen KPIs", "📋 Datos"])

    with tab_kpi:
        k = calcular_kpis(fdf, fc, dias)
        claves = list(k.keys())
        for i in range(0, len(claves), 4):
            for col, clave in zip(st.columns(4), claves[i:i + 4]):
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

        por_corte = kilos_por_corte(fdf, fc, dias)
        por_corte["Corte (cm)"] = por_corte["Corte"].map(corte_txt)
        orden = por_corte["Corte (cm)"].tolist()  # orden numérico, no alfabético

        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**Kilos por producto**")
            barras(fdf.assign(Producto=fdf["Producto"].replace("", SIN_DATO))
                   .groupby("Producto", as_index=False)["Kilos"].sum(), "Producto", "Kilos")
        with g2:
            st.markdown("**Kilos por corte (cm)**")
            barras(por_corte, "Corte (cm)", "Kilos", orden=orden)

        st.markdown(f"**Kilos por corte (cm) con más de {dias} días en stock**")
        viejos = por_corte[por_corte["Kilos_antiguos"] > 0]
        if viejos.empty:
            st.info(f"No hay stock con más de {dias} días para los filtros elegidos.")
        else:
            h1, h2 = st.columns(2)
            with h1:
                barras(viejos, "Corte (cm)", "Kilos_antiguos", titulo_y=f"Kilos > {dias} días", orden=orden)
            with h2:
                tabla = viejos[["Corte (cm)", "Unidades", "Kilos", "Kilos_antiguos", "% antiguo"]].rename(
                    columns={"Kilos_antiguos": f"Kilos > {dias} días", "% antiguo": "% de sus kilos"})
                st.dataframe(tabla, hide_index=True, column_config={
                    "% de sus kilos": st.column_config.NumberColumn(format="percent")})

    with tab_datos:
        st.dataframe(fdf, hide_index=True,
                     column_config={"Fecha": st.column_config.DateColumn(format="DD/MM/YYYY")})

    base = nombre.rsplit(".", 1)[0]
    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button("⬇️ Descargar Excel", data=_excel(fdf, tipo, fc, dias),
                       file_name=f"{base}.xlsx", key=f"xl_{nombre}",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇️ Descargar CSV", data=generar_csv(fdf),
                       file_name=f"{base}.csv", mime="text/csv", key=f"csv_{nombre}")
