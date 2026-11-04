import re
import openpyxl
import streamlit as st

# Configuración de la página web de Streamlit
st.set_page_config(
    page_title='Validador de Archivos Excel',
    page_icon='📊',
    layout='wide',
)


def validar_excel(file):
  wb = openpyxl.load_workbook(file, data_only=True)
  ws = wb.active

  total_leidos = 0
  registros_correctos = 0
  registros_con_errores = 0
  filas_con_error_nros = []

  # Contadores de errores por columna
  errores_por_columna = {
      'Col. A (Documento)': 0,
      'Col. B (RIF)': 0,
      'Col. C (Cuenta)': 0,
      'Col. D (Agencia)': 0,
      'Col. E (Monto)': 0,
      'Col. F (Canal)': 0,
      'Col. G (Estado)': 0,
  }

  detalle_errores_por_fila = []

  # Regex de validación
  patron_col_a = re.compile(r'^[VERNRvernr]\d{7,8}$')
  patron_col_b = re.compile(r'^[VERNRvernr]\d{7,9}$')
  patron_col_d = re.compile(r'^\d+\s*-\s*.+$')

  for num_fila, fila in enumerate(
      ws.iter_rows(min_row=1, max_row=ws.max_row), start=1
  ):
    # Omitir filas completamente vacías
    if all(cell.value is None for cell in fila):
      continue

    total_leidos += 1
    errores_en_esta_fila = []

    # Asegurar que existan al menos 7 columnas (A a G)
    while len(fila) < 7:

      class DummyCell:
        value = None
        data_type = None

      fila = list(fila) + [DummyCell()]

    # ----------------------------------------------------
    # COLUMNA A: Documento
    # ----------------------------------------------------
    val_a = str(fila[0].value).strip() if fila[0].value is not None else ''
    if not patron_col_a.match(val_a):
      errores_por_columna['Col. A (Documento)'] += 1
      errores_en_esta_fila.append(
          f'Columna A (Documento inválido: "{val_a}")'
      )

    # ----------------------------------------------------
    # COLUMNA B: RIF
    # ----------------------------------------------------
    val_b = str(fila[1].value).strip() if fila[1].value is not None else ''
    if not patron_col_b.match(val_b):
      errores_por_columna['Col. B (RIF)'] += 1
      errores_en_esta_fila.append(f'Columna B (RIF inválido: "{val_b}")')

    # ----------------------------------------------------
    # COLUMNA C: Cuenta (20 dígitos, 0138, apóstrofe en fx)
    # ----------------------------------------------------
    celda_c = fila[2]
    val_c = str(celda_c.value).strip() if celda_c.value is not None else ''
    tiene_apo_c = getattr(celda_c, 'quotePrefix', False)

    if (
        not val_c.isdigit()
        or len(val_c) != 20
        or not val_c.startswith('0138')
        or not tiene_apo_c
    ):
      errores_por_columna['Col. C (Cuenta)'] += 1
      motivo = 'Cuenta inválida/sin apóstrofe'
      if not tiene_apo_c:
        motivo = 'Falta apóstrofe en celda C'
      elif len(val_c) != 20:
        motivo = f'Longitud incorrecta ({len(val_c)} dígitos en vez de 20)'
      elif not val_c.startswith('0138'):
        motivo = 'No empieza con 0138'
      errores_en_esta_fila.append(f'Columna C ({motivo}: "{val_c}")')

    # ----------------------------------------------------
    # COLUMNA D: Agencia
    # ----------------------------------------------------
    val_d = str(fila[3].value).strip() if fila[3].value is not None else ''
    if not patron_col_d.match(val_d):
      errores_por_columna['Col. D (Agencia)'] += 1
      errores_en_esta_fila.append(f'Columna D (Agencia inválida: "{val_d}")')

    # ----------------------------------------------------
    # COLUMNA E: Monto (Punto decimal, apóstrofe en fx)
    # ----------------------------------------------------
    celda_e = fila[4]
    val_e_raw = celda_e.value
    tiene_apo_e = getattr(celda_e, 'quotePrefix', False)
    val_e_str = str(val_e_raw).strip() if val_e_raw is not None else ''

    es_valido_monto = False
    try:
      if val_e_raw is not None and not isinstance(val_e_raw, str):
        if '.' in str(val_e_raw) or isinstance(val_e_raw, (int, float)):
          es_valido_monto = True
      elif val_e_raw is not None:
        if re.match(r'^\d+(\.\d+)?$', val_e_str):
          es_valido_monto = True
    except Exception:
      es_valido_monto = False

    if not es_valido_monto or not tiene_apo_e:
      errores_por_columna['Col. E (Monto)'] += 1
      motivo_e = (
          'Falta apóstrofe o formato de monto inválido (use punto decimal)'
      )
      errores_en_esta_fila.append(f'Columna E ({motivo_e}: "{val_e_str}")')

    # ----------------------------------------------------
    # COLUMNA F: Canal (Debe ser 2072)
    # ----------------------------------------------------
    val_f = str(fila[5].value).strip() if fila[5].value is not None else ''
    if val_f != '2072':
      errores_por_columna['Col. F (Canal)'] += 1
      errores_en_esta_fila.append(
          f'Columna F (Canal inválido, debe ser 2072: "{val_f}")'
      )

    # ----------------------------------------------------
    # COLUMNA G: Estado (Debe ser N)
    # ----------------------------------------------------
    val_g = str(fila[6].value).strip() if fila[6].value is not None else ''
    if val_g.upper() != 'N':
      errores_por_columna['Col. G (Estado)'] += 1
      errores_en_esta_fila.append(
          f'Columna G (Estado inválido, debe ser N: "{val_g}")'
      )

    # Evaluar si la fila tiene errores
    if errores_en_esta_fila:
      registros_con_errores += 1
      filas_con_error_nros.append(num_fila)
      detalle_errores_por_fila.append({'fila': num_fila, 'errores': errores_en_esta_fila})
    else:
      registros_correctos += 1

  return {
      'total': total_leidos,
      'correctos': registros_correctos,
      'con_errores': registros_con_errores,
      'filas_error': filas_con_error_nros,
      'por_columna': errores_por_columna,
      'detalle_filas': detalle_errores_por_fila,
      'workbook': wb,
  }


def generar_excel_reparado(wb):
  ws = wb.active
  for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
    if row[2].value is not None:
      row[2].value = str(row[2].value).strip()
      row[2].quotePrefix = True

    if row[4].value is not None:
      val_e = str(row[4].value).replace(',', '.')
      row[4].value = val_e
      row[4].quotePrefix = True

    if len(row) >= 6:
      row[5].value = '2072'
    if len(row) >= 7:
      row[6].value = 'N'

  import io

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output


# --- DISEÑO DE LA INTERFAZ CON STREAMLIT ---
st.title('📊 Validador de Resultados y Celdas')
st.markdown(
    'Sube tu archivo Excel para validar automáticamente todas las reglas de'
    ' estructura, formato y apóstrofe.'
)

uploaded_file = st.file_uploader(
    'Selecciona o arrastra tu archivo Excel aquí', type=['xlsx', 'xls']
)

if uploaded_file is not None:
  with st.spinner('Analizando archivo...'):
    resultado = validar_excel(uploaded_file)

  # Tarjetas superiores
  col1, col2, col3, col4 = st.columns(4)

  with col1:
    st.metric(label='TOTAL FILAS PROCESADAS', value=resultado['total'])
  with col2:
    st.metric(label='REGISTROS CORRECTOS', value=resultado['correctos'])
  with col3:
    st.metric(label='REGISTROS CON ERRORES', value=resultado['con_errores'])
  with col4:
    filas_str = (
        ', '.join([f'Fila {f}' for f in resultado['filas_error']])
        if resultado['filas_error']
        else 'Ninguna'
    )
    st.metric(label='FILAS CON ERRORES (NROS)', value=filas_str)

  st.markdown('---')
  st.subheader('📋 Desglose de Errores por Columna:')

  cols_inf = st.columns(7)
  keys_cols = list(resultado['por_columna'].keys())

  for i, col_name in enumerate(keys_cols):
    with cols_inf[i]:
      cant_err = resultado['por_columna'][col_name]
      st.metric(label=col_name, value=cant_err)

  st.markdown('---')

  # --- NUEVA SECCIÓN: LISTADO DETALLADO DE FILAS CON ERRORES ---
  st.subheader('🔍 Detalle Fila por Fila de los Errores Encontrados')

  if resultado['detalle_filas']:
    for item in resultado['detalle_filas']:
      with st.expander(f"❌ Fila #{item['fila']} (Errores detectados: {len(item['errores'])})"):
        for err in item['errores']:
          st.markdown(f'- {err}')
  else:
    st.success('¡Excelente! No se encontraron errores en ninguna fila del archivo.')

  st.markdown('---')

  # Botón de descarga
  excel_limpio = generar_excel_reparado(resultado['workbook'])
  st.download_button(
      label='✨ Reparar y Descargar Excel Limpio',
      data=excel_limpio,
      file_name='archivo_reparado_limpio.xlsx',
      mime=(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ),
  )