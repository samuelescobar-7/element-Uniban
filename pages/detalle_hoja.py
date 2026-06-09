import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Detalle de la hoja")

hoja = st.session_state.get("detalle_hoja")
detalle_df = st.session_state.get("detalle_df")
detalle_df_k = st.session_state.get("detalle_df_k")

if hoja is None or (detalle_df is None and detalle_df_k is None):
    st.warning("No hay datos")
    st.stop()

st.subheader(f"Hoja: {hoja}")


# =========================
# HELPERS
# =========================
def df_to_excel_bytes(dfs: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def _orden_requerimientos(lista_dfs):
    """Devuelve los requerimientos en el orden original del archivo fuente (por columna Fila)."""
    df_ref = lista_dfs[0].sort_values("Fila")
    return list(dict.fromkeys(df_ref["Requerimiento"].tolist()))


def _pivot_ordenado(df_datos, col_valor, index_col, orden):
    """Hace pivot_table y reordena el índice según el orden original del archivo fuente."""
    pivot = (
        df_datos.pivot_table(
            index=index_col,
            columns="Proveedor",
            values=col_valor,
            aggfunc="first"
        )
        .fillna(0)
        .reindex(orden)
        .reset_index()
    )
    pivot.columns.name = None
    return pivot


# =========================
# TABLA DETALLE CUMPLIMIENTO
# =========================
if detalle_df is not None:
    st.markdown("#### Cumplimiento por requerimiento")

    df = detalle_df.copy()
    df["Cumplimiento_%"] = df["Peso_Total"] * 100

    # Orden original: usar columna "Fila" del primer proveedor
    _proveedores = df["Proveedor"].unique().tolist()
    _lista_dfs = [df[df["Proveedor"] == p] for p in _proveedores if not df[df["Proveedor"] == p].empty]
    _orden_reqs = _orden_requerimientos(_lista_dfs)

    df_pivot = _pivot_ordenado(df, "Cumplimiento_%", "Requerimiento", _orden_reqs)

    df_pivot_fmt = df_pivot.copy()
    for col in df_pivot_fmt.columns:
        if col != "Requerimiento":
            df_pivot_fmt[col] = df_pivot_fmt[col].apply(lambda x: f"{x:.2f}%")

    st.dataframe(df_pivot_fmt, use_container_width=True)
    st.download_button(
        "⬇️ Descargar cumplimiento por requerimiento",
        df_to_excel_bytes({"Cumplimiento por requerimiento": df_pivot}),
        file_name=f"cumplimiento_requerimiento_{hoja}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_pivot_cumplimiento"
    )

    resumen = detalle_df.groupby("Proveedor")["Peso_Total"].mean().mul(100).round(2).reset_index()
    resumen_fmt = resumen.copy()
    resumen_fmt["Cumplimiento_%"] = resumen_fmt["Peso_Total"].apply(lambda x: f"{x:.2f}%")
    resumen_fmt = resumen_fmt.drop(columns=["Peso_Total"])

    st.markdown("**Resumen cumplimiento**")
    st.dataframe(resumen_fmt)
    resumen_dl = resumen.rename(columns={"Peso_Total": "Cumplimiento_%"})
    st.download_button(
        "⬇️ Descargar resumen cumplimiento",
        df_to_excel_bytes({"Resumen cumplimiento": resumen_dl}),
        file_name=f"resumen_cumplimiento_{hoja}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_resumen_cumplimiento"
    )

# =========================
# TABLA DETALLE CALIDAD (K)
# =========================
if detalle_df_k is not None:
    st.markdown("#### Calidad por requerimiento")

    df_k = detalle_df_k.copy()
    df_k["Calidad_%"] = df_k["Peso_K"] * 100

    # Orden original: usar columna "Fila" del primer proveedor
    _proveedores_k = df_k["Proveedor"].unique().tolist()
    _lista_dfs_k = [df_k[df_k["Proveedor"] == p] for p in _proveedores_k if not df_k[df_k["Proveedor"] == p].empty]
    _orden_reqs_k = _orden_requerimientos(_lista_dfs_k)

    df_pivot_k = _pivot_ordenado(df_k, "Calidad_%", "Requerimiento", _orden_reqs_k)

    df_pivot_k_fmt = df_pivot_k.copy()
    for col in df_pivot_k_fmt.columns:
        if col != "Requerimiento":
            df_pivot_k_fmt[col] = df_pivot_k_fmt[col].apply(lambda x: f"{x:.2f}%")

    st.dataframe(df_pivot_k_fmt, use_container_width=True)
    st.download_button(
        "⬇️ Descargar calidad por requerimiento",
        df_to_excel_bytes({"Calidad por requerimiento": df_pivot_k}),
        file_name=f"calidad_requerimiento_{hoja}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_pivot_calidad"
    )

    resumen_k = detalle_df_k.groupby("Proveedor")["Peso_K"].mean().mul(100).round(2).reset_index()
    resumen_k_fmt = resumen_k.copy()
    resumen_k_fmt["Calidad_%"] = resumen_k_fmt["Peso_K"].apply(lambda x: f"{x:.2f}%")
    resumen_k_fmt = resumen_k_fmt.drop(columns=["Peso_K"])

    st.markdown("**Resumen calidad**")
    st.dataframe(resumen_k_fmt)
    resumen_k_dl = resumen_k.rename(columns={"Peso_K": "Calidad_%"})
    st.download_button(
        "⬇️ Descargar resumen calidad",
        df_to_excel_bytes({"Resumen calidad": resumen_k_dl}),
        file_name=f"resumen_calidad_{hoja}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_resumen_calidad"
    )

# =========================
# EXPORTAR EXCEL COMPLETO
# =========================
sheets_completo = {}
if detalle_df is not None:
    sheets_completo["Detalle cumplimiento"] = df_pivot
    sheets_completo["Resumen cumplimiento"] = resumen.rename(columns={"Peso_Total": "Cumplimiento_%"})
if detalle_df_k is not None:
    sheets_completo["Detalle calidad"] = df_pivot_k
    sheets_completo["Resumen calidad"] = resumen_k.rename(columns={"Peso_K": "Calidad_%"})

st.divider()
st.download_button(
    "⬇️ Descargar todo (Excel completo)",
    df_to_excel_bytes(sheets_completo),
    file_name=f"detalle_{hoja}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="dl_completo"
)

# =========================
# VOLVER
# =========================
if st.button("Volver"):
    st.session_state["detalle_df_k"] = None
    st.switch_page("app.py")