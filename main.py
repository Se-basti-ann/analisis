import os
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from fastapi import FastAPI, Form, UploadFile, File, HTTPException, Response
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from io import BytesIO
import logging
from collections import defaultdict
import re

app = FastAPI()
logger = logging.getLogger(__name__)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://main.d32bb122o9jw4d.amplifyapp.com"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def generate_resumen_tecnicos(writer, datos_combinados, df_originales):

    instalados_por_tecnico = {}
    retirados_por_tecnico = {}
    luminarias_por_tecnico = {}
    tecnicos = set()
    
    # Mapeo de nodos a técnicos
    nodo_tecnico_map = {}
    
    # Primero, construir un mapa de nodo -> técnico desde los DataFrames originales
    for df in df_originales.values():
        if "4.Nombre del Técnico Instalador" in df.columns and "1.NODO DEL POSTE." in df.columns:
            for _, fila in df.iterrows():
                nodo = str(fila["1.NODO DEL POSTE."]).strip()
                tecnico = str(fila.get("4.Nombre del Técnico Instalador", "Sin técnico")).strip()
                
                # Normalizar el nombre del técnico (eliminar espacios extra, capitalizar correctamente)
                tecnico = " ".join([palabra.capitalize() for palabra in tecnico.split() if palabra])
                
                if not tecnico or tecnico.lower() in ['na', 'n/a', 'ninguno']:
                    tecnico = "Sin técnico"
                
                # Manejo de nodos 0
                ot = fila["2.Nro de O.T."]
                if nodo in ['0', '0.0']:
                    nodo = f"0_{ot}"  # Usar el formato que se usa en procesar_archivo_modernizacion
                
                nodo_tecnico_map[nodo] = tecnico
                tecnicos.add(tecnico)
    
    # Ordenar técnicos para mantener consistencia
    tecnicos_ordenados = sorted(tecnicos)
    
    # Ahora procesar datos combinados
    for ot, info in datos_combinados.items():
        # Proceso para códigos N1
        for key, nodos_data in info.get('codigos_n1', {}).items():
            nombre = f"LUMINARIA {key}"
            
            for nodo, codigos in nodos_data.items():
                tecnico = nodo_tecnico_map.get(nodo.split('_')[0], "Sin técnico")
                
                if tecnico not in luminarias_por_tecnico:
                    luminarias_por_tecnico[tecnico] = {}
                
                if nombre not in luminarias_por_tecnico[tecnico]:
                    luminarias_por_tecnico[tecnico][nombre] = {
                        'tipo': 'LUMINARIA',
                        'unidad': 'UND',
                        'total': 0,
                        'ots': defaultdict(int)
                    }
                
                # Contar el número de códigos en este nodo
                codigos_count = len(codigos)
                luminarias_por_tecnico[tecnico][nombre]['total'] += codigos_count
                luminarias_por_tecnico[tecnico][nombre]['ots'][ot] += codigos_count
        
        # Proceso para códigos N2
        for key, nodos_data in info.get('codigos_n2', {}).items():
            nombre = f"LUMINARIA {key}"
            
            for nodo, codigos in nodos_data.items():
                tecnico = nodo_tecnico_map.get(nodo.split('_')[0], "Sin técnico")
                
                if tecnico not in luminarias_por_tecnico:
                    luminarias_por_tecnico[tecnico] = {}
                
                if nombre not in luminarias_por_tecnico[tecnico]:
                    luminarias_por_tecnico[tecnico][nombre] = {
                        'tipo': 'LUMINARIA',
                        'unidad': 'UND',
                        'total': 0,
                        'ots': defaultdict(int)
                    }
                
                # Contar el número de códigos en este nodo
                codigos_count = len(codigos)
                luminarias_por_tecnico[tecnico][nombre]['total'] += codigos_count
                luminarias_por_tecnico[tecnico][nombre]['ots'][ot] += codigos_count
            
        # Procesar materiales instalados
        for mat_key, nodo_quantities in info['materiales'].items():
            _, nombre = mat_key.split('|', 1)
            
            if nombre.strip().upper() == "NINGUNO":
                continue
            
            for nodo, cantidad in nodo_quantities.items():
                tecnico = nodo_tecnico_map.get(nodo.split('_')[0], "Sin técnico")
                
                if tecnico not in instalados_por_tecnico:
                    instalados_por_tecnico[tecnico] = {}
                
                if nombre not in instalados_por_tecnico[tecnico]:
                    instalados_por_tecnico[tecnico][nombre] = {
                        'tipo': 'INSTALADO',
                        'unidad': 'UND',
                        'total': 0,
                        'ots': defaultdict(int)
                    }
                instalados_por_tecnico[tecnico][nombre]['total'] += cantidad
                instalados_por_tecnico[tecnico][nombre]['ots'][ot] += cantidad
        
        # Procesar materiales retirados
        for mat_key, nodo_quantities in info.get('materiales_retirados', {}).items():
            _, nombre = mat_key.split('|', 1)
            
            if nombre.strip().upper() == "NINGUNO":
                continue
            
            for nodo, cantidad in nodo_quantities.items():
                tecnico = nodo_tecnico_map.get(nodo.split('_')[0], "Sin técnico")
                
                if tecnico not in retirados_por_tecnico:
                    retirados_por_tecnico[tecnico] = {}
                
                if nombre not in retirados_por_tecnico[tecnico]:
                    retirados_por_tecnico[tecnico][nombre] = {
                        'tipo': 'RETIRADO',
                        'unidad': 'UND',
                        'total': 0,
                        'ots': defaultdict(int)
                    }
                retirados_por_tecnico[tecnico][nombre]['total'] += cantidad
                retirados_por_tecnico[tecnico][nombre]['ots'][ot] += cantidad

    # Verificación de totales
    for tecnico, materiales in luminarias_por_tecnico.items():
        for nombre, info in materiales.items():
            # Recalcular el total sumando todas las cantidades por OT para confirmar
            total_from_ots = sum(info['ots'].values())
            if info['total'] != total_from_ots:
                # Ajustar si hay discrepancia
                info['total'] = total_from_ots
    
    # Crear la hoja para el resumen de técnicos
    sheet_name = 'Resumen_tecnicos'
    worksheet = writer.book.create_sheet(sheet_name)
    
    # Configuración de estilos
    header_font = Font(bold=True)
    section_font = Font(bold=True, size=12)
    section_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # Columnas del resumen
    columns = ['Nombre Material', 'Unidad', 'Cantidad Total'] + tecnicos_ordenados
    
    # Si no hay datos, salir
    if not (luminarias_por_tecnico or instalados_por_tecnico or retirados_por_tecnico):
        return
    
    # Escribir encabezados de columnas
    for col, header in enumerate(columns, 1):
        cell = worksheet.cell(row=1, column=col, value=header)
        cell.font = header_font
    
    current_row = 2
    
    # Función para escribir una sección
    def write_section(title, data_by_tecnico, row):
        # Encabezado de sección
        cell = worksheet.cell(row=row, column=1, value=title)
        cell.font = section_font
        cell.fill = section_fill
        cell.alignment = center_alignment
        
        # Combinar celdas para el encabezado
        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(columns))
        
        row += 1
        
        # Recolectar todos los materiales únicos para esta sección
        all_materials = set()
        for tecnico in tecnicos_ordenados:
            if tecnico in data_by_tecnico:
                all_materials.update(data_by_tecnico[tecnico].keys())
        
        # Escribir datos por material
        for material in sorted(all_materials):
            worksheet.cell(row=row, column=1, value=material)
            worksheet.cell(row=row, column=2, value='UND')
            
            # Calcular el total sumando de todos los técnicos
            total = sum(
                data_by_tecnico.get(tecnico, {}).get(material, {}).get('total', 0)
                for tecnico in tecnicos_ordenados
            )
            worksheet.cell(row=row, column=3, value=total)
            
            # Cantidades por técnico
            for i, tecnico in enumerate(tecnicos_ordenados, 4):
                cantidad = data_by_tecnico.get(tecnico, {}).get(material, {}).get('total', 0)
                worksheet.cell(row=row, column=i, value=cantidad)
            
            row += 1
            
        return row

    # Escribir sección de luminarias
    if luminarias_por_tecnico:
        current_row = write_section("LUMINARIAS INSTALADAS", luminarias_por_tecnico, current_row)
        current_row += 1  # Espacio entre secciones
    
    # Escribir sección de materiales instalados
    if instalados_por_tecnico:
        current_row = write_section("MATERIALES INSTALADOS", instalados_por_tecnico, current_row)
        current_row += 1  # Espacio entre secciones
    
    # Escribir sección de materiales retirados
    if retirados_por_tecnico:
        current_row = write_section("MATERIALES RETIRADOS", retirados_por_tecnico, current_row)
    
    # Ajustar anchos de columna
    for idx, col in enumerate(columns, 1):
        col_width = max(len(str(col)), 15) + 2
        # Para la primera columna, usar un ancho mayor
        if idx == 1:
            col_width = 40
        worksheet.column_dimensions[get_column_letter(idx)].width = col_width
    
    # Configurar congelación de paneles
    worksheet.freeze_panes = 'D2'

def normalizar_barrio(barrio):
    barrio_str = str(barrio).strip()
    try:
        pd.to_datetime(barrio_str, dayfirts=True)
        return 'Sin barrio'
    except Exception:
        pass
    if barrio_str in ['0', '0.0']:
        return 'Sin barrio'
    return barrio_str.title().strip()
           
def procesar_archivo_modernizacion(file: UploadFile):
    try:
        contenido = file.file.read()
        xls = pd.ExcelFile(BytesIO(contenido))
        
        datos = defaultdict(lambda: {
            'nodos': [],
            'nodo_counts': defaultdict(int),
            'codigos_n1': defaultdict(lambda: defaultdict(set)),
            'codigos_n2': defaultdict(lambda: defaultdict(set)),
            'materiales': defaultdict(lambda: defaultdict(int)),    
            'materiales_retirados': defaultdict(lambda: defaultdict(int)),
            'aspectos_materiales': defaultdict(lambda: defaultdict(lambda: set())),
            'aspectos_retirados': defaultdict(lambda: defaultdict(lambda: set())),
            'fechas_sync': defaultdict(str)  # Nuevo campo para almacenar fechas de sincronización
        })                 
        
        datos_por_barrio = defaultdict(lambda: {
            'materiales_instalados': defaultdict(lambda: defaultdict(int)),
            'materiales_retirados': defaultdict(lambda: defaultdict(int))
                                                })
        dfs_originales = {}
        counter_0 = defaultdict(int)
        # Configurar columnas         
        COL_INICIO = "BH"
        COL_FIN = "BO"
        idx_inicio = column_index_from_string(COL_INICIO) - 1
        idx_fin = column_index_from_string(COL_FIN) - 1
        
        # Diccionario para normalizar nodos
        nodos_normalizados = {}
        
        for hoja in xls.sheet_names:
            df = xls.parse(hoja)
            
            # Validar columnas requeridas
            required_columns = {
                "2.Nro de O.T.", "1.NODO DEL POSTE.",
                "2.CODIGO DE LUMINARIA INSTALADA N1.", "3.POTENCIA DE LUMINARIA INSTALADA (W)",
                "6.CODIGO DE LUMINARIA INSTALADA N2.", "7.POTENCIA DE LUMINARIA INSTALADA (W)",
                "1. Describa Aspectos que Considere se deben tener en cuenta.",
                "FechaSincronizacion"
            }
            if not required_columns.issubset(df.columns):
                continue          
            
            dfs_originales[hoja] = df
            # Parsear y ordenar por FechaSincronizacion
            df['FechaSincronizacion'] = (
                df['FechaSincronizacion']
                .astype(str)
                .str.replace('a. m.', 'AM', case=False)
                .str.replace('p. m.', 'PM', case=False)
            )
            df['FechaSincronizacion'] = pd.to_datetime(
                df['FechaSincronizacion'],
                format='%d/%m/%Y %I:%M:%S %p',
                errors='coerce'
            )
            df = df.sort_values(by='FechaSincronizacion', ascending=True)
            # PROCESAR MATERIALES RETIRADOS
            pattern_codigo = re.compile(r'^\d+\.CODIGO DE (LUMINARIA|BOMBILLA|FOTOCELDA) RETIRADA (N\d+)\.?$', re.IGNORECASE)
            pattern_potencia = re.compile(r'^\d+\.POTENCIA DE (LUMINARIA|BOMBILLA) RETIRADA (N\d+)\.?\(W\)$', re.IGNORECASE)
            
            codigo_columns = {}
            potencia_columns = {}
            
            for col in df.columns:
                col_str = str(col).strip()
                # PORECESAR CODIGOS RETIRADA
                codigo_match = pattern_codigo.match(col_str)
                if codigo_match:
                    tipo = codigo_match.group(1).upper()
                    n = codigo_match.group(2).upper()
                    codigo_columns[(tipo, n)] = col_str
                else:
                    # PROCESAR COLUMNAS POTENCIA
                    potencia_match = pattern_potencia.match(col_str)
                    if potencia_match:
                        tipo = potencia_match.group(1).upper()
                        n = potencia_match.group(2).upper()
                        potencia_columns[(tipo, n)] = col_str
            
            # ========== PROCESAR MATERIALES ORIGINALES (MATERIAL X - CANTIDAD X) ==========
            material_cols = [col for col in df.columns if re.match(r'^(MATERIAL|Material)\s\d+$', col)]
            cantidad_cols = [col for col in df.columns if re.match(r'^(CANTIDAD MATERIAL|CANTIDAD DE MATERIAL)\s\d+$', col)]
            
            columnas_bh_bo = []
            if len(df.columns) > idx_fin:
                columnas_bh_bo = df.columns[idx_inicio : idx_fin + 1]
                        
            for _, fila in df.iterrows():
                ot = fila["2.Nro de O.T."]   
                
                barrio = fila.get("3.Barrio", "")
                barrio_normalizado = normalizar_barrio(barrio)
                                     
                original_nodo = str(fila["1.NODO DEL POSTE."]).strip().replace(' ', '').replace('-', '').upper()
                
                # Guardar la fecha de sincronización para este nodo/registro
                fecha_sincronizacion = fila["FechaSincronizacion"]
    
                # Normalizar nodos con valor 0 o similares
                if original_nodo in ['0', '0.0', '0.00', 'nan', 'NaN', 'None', '']:
                    # Usar un formato consistente para todos los nodos "0"
                    counter_0[ot] += 1
                    nodo = f"0_{ot}_{counter_0[ot]}"  # Formato: 0_OT_contador
                else:
                    # Normalizar el formato del nodo para evitar duplicados por formato
                    nodo_normalizado = original_nodo.replace('.0', '').replace('.00', '')
                    
                    # Verificar si este nodo ya ha sido normalizado antes
                    if (ot, nodo_normalizado) in nodos_normalizados:
                        nodo = nodos_normalizados[(ot, nodo_normalizado)]
                    else:
                        count = datos[ot]['nodo_counts'][nodo_normalizado] + 1
                        datos[ot]['nodo_counts'][nodo_normalizado] = count
                        nodo = f"{nodo_normalizado}_{count}" if count > 1 else nodo_normalizado
                        nodos_normalizados[(ot, nodo_normalizado)] = nodo
                
                datos[ot]['nodos'].append(nodo)
                
                # Guardar la fecha de sincronización para este nodo
                if pd.notna(fecha_sincronizacion):
                    datos[ot]['fechas_sync'][nodo] = fecha_sincronizacion.strftime('%d/%m/%Y %H:%M:%S')
                else:
                    datos[ot]['fechas_sync'][nodo] = "Sin fecha"
                                    
                codigo_n1 = fila["2.CODIGO DE LUMINARIA INSTALADA N1."]
                potencia_n1 = fila["3.POTENCIA DE LUMINARIA INSTALADA (W)"]
                
                codigo_n2 = fila["6.CODIGO DE LUMINARIA INSTALADA N2."]
                potencia_n2 = fila["7.POTENCIA DE LUMINARIA INSTALADA (W)"]
                
                for col in columnas_bh_bo:
                    cantidad = fila[col]
                    # Limpiar NaN y convertir a 0
                    if pd.isna(cantidad):
                        cantidad = 0
                    else:
                        cantidad = int(cantidad)  # Asegurar entero
                    #if pd.notna(cantidad) and float(cantidad) > 0:
                    if cantidad > 0:
                        nombre_material = str(col).split('.', 1)[-1].strip().upper()
                        key = f"MATERIAL_RETIRADO|{nombre_material}"
                        datos[ot]['materiales_retirados'][key][nodo] += cantidad
                
                        # Capture aspect here
                        aspecto = fila["1. Describa Aspectos que Considere se deben tener en cuenta."]
                        if pd.notna(aspecto):
                            aspecto_limpio = str(aspecto).strip().upper()
                            if aspecto_limpio not in ['', 'NA', 'NINGUNO', 'N/A']:
                                datos[ot]['aspectos_retirados'][key][nodo].add(aspecto_limpio)
                    
                # Procesar códigos N1 y N2 para que sean accesibles para las funciones de mano de obra
                if pd.notna(codigo_n1) and str(codigo_n1).strip() not in ['', '0', '0.0'] or pd.notna(potencia_n1) and str(potencia_n1) != "0" and float(potencia_n1) != 0:
                    try:
                        potencia_val = float(potencia_n1)
                        if potencia_val == 0:
                            key = "CODIGO 1 LUMINARIA INSTALADA"
                        else:
                            if potencia_val.is_integer():
                                key = f"CODIGO 1 LUMINARIA INSTALADA {int(potencia_val)} W"
                            else:
                                key = f"CODIGO 1 LUMINARIA INSTALADA {potencia_val} W"
                        datos[ot]['codigos_n1'][key][nodo].add(str(codigo_n1).strip().upper())
                    except:
                        pass                          
                
                # Agregar como material para que sea visible para las funciones de mano de obra
                if pd.notna(codigo_n1) and str(codigo_n1).strip() not in ['', '0', '0.0']:
                    material_key = f"MATERIAL|CODIGO DE LUMINARIA INSTALADA N1"
                    datos[ot]['materiales'][material_key][nodo] += 1  # Cada código cuenta como 1 unidad
                
                if pd.notna(codigo_n2) and str(codigo_n2).strip() not in ['', '0', '0.0'] or pd.notna(potencia_n2) and str(potencia_n2) != "0" and float(potencia_n2) != 0:
                    try:
                        potencia_val = float(potencia_n2)
                        if potencia_val == 0:
                            key = "CODIGO 2 LUMINARIA INSTALADA"
                        else:
                            if potencia_val.is_integer():
                                key = f"CODIGO 2 LUMINARIA INSTALADA {int(potencia_val)} W"
                            else:
                                key = f"CODIGO 2 LUMINARIA INSTALADA {potencia_val} W"
                        datos[ot]['codigos_n2'][key][nodo].add(str(codigo_n2).strip().upper())
                    except:
                        pass
                                            
                # Agregar como material para que sea visible para las funciones de mano de obra
                if pd.notna(codigo_n2) and str(codigo_n2).strip() not in ['', '0', '0.0']:
                    material_key = f"MATERIAL|CODIGO DE LUMINARIA INSTALADA N2"
                    datos[ot]['materiales'][material_key][nodo] += 1
                
                # PROCESAR MATERIALES RETIRADOS
                for (tipo, n), col_codigo in codigo_columns.items():
                    codigo_val = fila.get(col_codigo)
                    col_potencia = None
                    potencia_val = None

                    if tipo in ['LUMINARIA', 'BOMBILLA']:
                        col_potencia = potencia_columns.get((tipo, n), None)
                        if col_potencia:
                            potencia_val = fila.get(col_potencia)
                    
                    codigo_es_valido = (
                        pd.notna(codigo_val)
                        and str(codigo_val).strip() not in ['', '0', '0.0']
                    )
                    
                    potencia_es_valida = False
                    if col_potencia and pd.notna(potencia_val):
                        try: 
                            potencia_float = float(potencia_val)
                            potencia_es_valida = potencia_float > 0
                        except:
                            pass
                    
                    if codigo_es_valido or potencia_es_valida:
                        if tipo == 'FOTOCELDA':
                            entry_name = f"FOTOCELDA RETIRADA {n}"
                        else:
                            potencia_str = ""
                            if potencia_es_valida:
                                potencia_str = (
                                    f"{int(potencia_float)}W"
                                    if potencia_float.is_integer()
                                    else f"{potencia_float}W"
                                )                                               
    
                            if potencia_str:
                                entry_name = f"{tipo} RETIRADA {n} {potencia_str}".strip()
                            else:
                                entry_name = f"{tipo} RETIRADA {n}".strip()
                                
                        key = f"MATERIAL_RETIRADO|{entry_name}"
                        datos[ot]['materiales_retirados'][key][nodo] += 1
                        aspecto = fila["1. Describa Aspectos que Considere se deben tener en cuenta."]
                        if pd.notna(aspecto):
                            aspecto_limpio = str(aspecto).strip().upper()
                            if aspecto_limpio not in ['', 'NA', 'NINGUNO', 'N/A']:  # Filtrar valores no válidos
                                datos[ot]['aspectos_retirados'][key][nodo].add(aspecto_limpio)  
                
                # Procesar materiales INSTALADOS
                for mat_col, cant_col in zip(material_cols, cantidad_cols):
                    material = fila[mat_col]
                    cantidad = fila[cant_col]
                    if pd.notna(material) and pd.notna(cantidad) and float(cantidad) > 0.0:
                        key = f"MATERIAL|{material}".strip().upper()
                        datos[ot]['materiales'][key][nodo] += cantidad  
                        datos_por_barrio[barrio_normalizado]['materiales_instalados'][key][ot] += cantidad
                        aspecto = fila["1. Describa Aspectos que Considere se deben tener en cuenta."]
                        if pd.notna(aspecto):
                            aspecto_limpio = str(aspecto).strip().upper()
                            if aspecto_limpio not in ['', 'NA', 'NINGUNO', 'N/A']:  # Filtrar valores no válidos
                                datos[ot]['aspectos_materiales'][key][nodo].add(aspecto_limpio)                               
                                               
        return datos, datos_por_barrio, dfs_originales

    except Exception as e:
        logger.error(f"Error procesando {file.filename}: {str(e)}")
        raise HTTPException(500, detail=f"Error en archivo {file.filename}")

def procesar_archivo_mantenimiento(file: UploadFile):
    try:
        contenido = file.file.read()
        xls = pd.ExcelFile(BytesIO(contenido))
        datos = defaultdict(lambda: {
            'nodos': set(),
            'materiales': defaultdict(lambda: defaultdict(int))
        })

        material_pattern = re.compile(r'^MATERIAL\s\d+$', re.IGNORECASE)
        cantidad_pattern = re.compile(r'^CANTIDAD MATERIAL\s\d+$', re.IGNORECASE)

        for hoja in xls.sheet_names:
            df = xls.parse(hoja).rename(columns=lambda x: str(x).strip())
            required_columns = {"6.Nro.Orden Energis", "5.Nodo"}
            if not required_columns.issubset(df.columns):
                continue

            df["5.Nodo"] = df["5.Nodo"].astype(str)
            ot_nodos = df[["6.Nro.Orden Energis", "5.Nodo"]].drop_duplicates()
            for ot, nodo in ot_nodos.itertuples(index=False, name=None):
                datos[ot]['nodos'].add(nodo)

            material_cols = [col for col in df.columns if material_pattern.match(col)]
            cantidad_cols = [col for col in df.columns if cantidad_pattern.match(col)]
            paired_columns = []
            material_nums = {re.search(r'\d+', col).group() for col in material_cols}
            for col in cantidad_cols:
                num = re.search(r'\d+', col).group()
                if num in material_nums:
                    mat_col = next(c for c in material_cols if num in c)
                    paired_columns.append((mat_col, col))

            materiales_data = []
            for mat_col, cant_col in paired_columns:
                temp_df = df[["6.Nro.Orden Energis", "5.Nodo", mat_col, cant_col]].copy()
                temp_df.columns = ["OT", "Nodo", "Material", "Cantidad"]
                materiales_data.append(temp_df)

            combined_df = pd.concat(materiales_data, ignore_index=True)
            combined_df = combined_df[
                (combined_df['Material'].notna()) &
                (~combined_df['Material'].str.strip().str.upper().isin(['', 'NINGUNO', 'NA'])) &
                (pd.to_numeric(combined_df['Cantidad'], errors='coerce') > 0)
            ]
            
            if combined_df.empty:
                continue

            combined_df['Material'] = combined_df['Material'].str.strip().str.upper()
            combined_df['Cantidad'] = pd.to_numeric(combined_df['Cantidad'])
            grouped = combined_df.groupby(['OT', 'Nodo', 'Material'])['Cantidad'].sum().reset_index()

            for row in grouped.itertuples(index=False):
                ot, nodo, material, cantidad = row
                key = f"MATERIAL|{material}"
                datos[ot]['materiales'][key][nodo] += cantidad

        return datos

    except Exception as e:
        logger.error(f"Error procesando {file.filename}: {str(e)}")
        raise HTTPException(500, detail=f"Error en archivo {file.filename}")

def agregar_hoja_asociaciones(writer, datos_combinados):
    wb = writer.book
    ws = wb.create_sheet("Asociaciones")

    # Estilos
    thin   = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    red    = PatternFill("solid", fgColor="FF0000")
    gray   = PatternFill("solid", fgColor="CCCCCC")
    bold   = Font(bold=True)
    center = Alignment(horizontal='center', vertical='center')

    # Cargo la plantilla (una sola vez)
    plantilla = cargar_plantilla_mano_obra()

    # Los cuatro bloques con sus keywords
    bloques = [
        ("Postes",
         ["APERTURA", "APLOMADA", "CONCRETADA", "HINCADA"],
         ["POSTE"]),
        ("Instalación luminarias",
         ["INSTALACION LUMINARIAS"],
         ["LUMINARIA", "FOTOCELDA", "GRILLETE", "BRAZO"]),
        ("Conexión a tierra",
         ["CONEXIÓN A CABLE A TIERRA", "INSTALACION KIT SPT", "INSTALACION DE ATERRIZAJES"],
         ["KIT DE PUESTA A TIERRA", "CONECT PERF", "CONECTOR BIME/COM", "ALAMBRE", "TUERCA", "TORNILLO", "VARILLA"]),
        ("Desmontaje / Transporte",
         ["DESMONTAJE", "TRANSPORTE", "TRANSP."],
         ["ALAMBRE", "BRAZO", "CÓDIGO", "CABLE"]),
        ("Instalación de cables",
         ["INSTALACION CABLE"],
         ["CABLE", "TPX", "ALAMBRE"]),
        ("Otros trabajos",
         ["VESTIDA", "CAJA", "PINTADA", "EXCAVACION", "RECUPERACION", "SOLDADURA", "INSTALACION TRAMA", "INSTALACION CORAZA"],
         ["PERCHA", "CAJA", "TUBO", "TUBERIA", "CONDUIT"])
    ]

    row = 1
    for ot, info in datos_combinados.items():
        # 1) Cabecera de OT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        c = ws.cell(row=row, column=1, value=f"OT: {ot}")
        c.font = Font(bold=True, size=14)
        c.alignment = center
        row += 2

        # 2) Defino nodos ordenados por fecha de sincronización
        nodos = sorted(
            info.get('fechas_sync', {}).keys(),
            key=lambda n: pd.to_datetime(
                info['fechas_sync'].get(n, ""),
                format='%d/%m/%Y %H:%M:%S',
                errors='coerce'
            )
        )

        # 3) Por cada nodo...
        for nodo in nodos:
            # Cabecera de nodo
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
            cn = ws.cell(row=row, column=1, value=f"Nodo: {nodo}")
            cn.font = Font(bold=True, size=12)
            cn.fill = gray
            row += 1

            # 4) Los bloques
            for titulo, kw_mo, kw_mat in bloques:
                # Encabezado de bloque
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
                cb = ws.cell(row=row, column=1, value=titulo)
                cb.font = Font(bold=True, color="FFFFFF")
                cb.fill = red
                row += 1

                # Cabeceras de mini-tabla
                headers = [
                    "Mano de obra", "Unidad", "Cant. MO",
                    "Material Nuevo", "Cant. Mat.",
                    "Material Ret.",   "Cant. Mat."
                ]
                for col_idx, h in enumerate(headers, start=1):
                    ch = ws.cell(row=row, column=col_idx, value=h)
                    ch.font = bold
                    ch.border = border
                    ch.alignment = center
                    if col_idx == 1:
                        ch.fill = red
                row += 1

                # Filtrar partidas de la plantilla que pertenecen a este bloque
                partidas = [
                    item for item in plantilla
                    if any(k in item['DESCRIPCION MANO DE OBRA'].upper() for k in kw_mo)
                ]

                # Para cada partida, calcular la mano de obra necesaria
                for item in partidas:
                    desc = item['DESCRIPCION MANO DE OBRA']
                    und = item['UNIDAD']
                    
                    # Usar la función mejorada para calcular cantidades de mano de obra
                    total_mo, lst_inst, lst_ret = calcular_cantidad_mano_obra(
                        desc, 
                        info.get('materiales', {}), 
                        info.get('materiales_retirados', {}),
                        nodo
                    )
                    
                    # Si no hay mano de obra requerida, continuamos con la siguiente partida
                    if total_mo == 0:
                        continue
                    
                    # Contar materiales instalados y retirados
                    # Para materiales instalados, obtener las cantidades reales
                    materiales_inst_con_qty = []
                    for mat_name in lst_inst:
                        for material_key, nodos_qty in info.get('materiales', {}).items():
                            if material_key.split("|")[1].upper() == mat_name and nodo in nodos_qty:
                                qty = nodos_qty[nodo]
                                materiales_inst_con_qty.append(f"{mat_name} ({qty})")
                                break
                        else:
                            materiales_inst_con_qty.append(mat_name)
                    
                    # Para materiales retirados, obtener las cantidades reales
                    materiales_ret_con_qty = []
                    for mat_name in lst_ret:
                        for material_key, nodos_qty in info.get('materiales_retirados', {}).items():
                            if material_key.split("|")[1].upper() == mat_name and nodo in nodos_qty:
                                qty = nodos_qty[nodo]
                                materiales_ret_con_qty.append(f"{mat_name} ({qty})")
                                break
                        else:
                            materiales_ret_con_qty.append(mat_name)
                    
                    # Calcular sumas totales de materiales
                    sum_inst = sum(
                        nodos_qty[nodo] 
                        for material_key, nodos_qty in info.get('materiales', {}).items() 
                        if material_key.split("|")[1].upper() in lst_inst and nodo in nodos_qty
                    ) if lst_inst else 0
                    
                    sum_ret = sum(
                        nodos_qty[nodo] 
                        for material_key, nodos_qty in info.get('materiales_retirados', {}).items() 
                        if material_key.split("|")[1].upper() in lst_ret and nodo in nodos_qty
                    ) if lst_ret else 0

                    # Escribo la fila
                    ws.cell(row=row, column=1, value=desc).border = border
                    ws.cell(row=row, column=2, value=und).border = border
                    ws.cell(row=row, column=3, value=total_mo).border = border

                    ws.cell(row=row, column=4, value="; ".join(materiales_inst_con_qty)).border = border
                    ws.cell(row=row, column=5, value=sum_inst if sum_inst > 0 else "").border = border
                    
                    ws.cell(row=row, column=6, value="; ".join(materiales_ret_con_qty)).border = border
                    ws.cell(row=row, column=7, value=sum_ret if sum_ret > 0 else "").border = border
                    row += 1

                row += 1  # espacio tras bloque

            row += 2  # espacio tras nodo

        row += 4  # espacio tras OT

    # Ajusto ancho de columnas
    for col_idx in range(1, 8):
        ws.column_dimensions[get_column_letter(col_idx)].width = 30
        
def generate_resumen_general(writer, datos_combinados):

    # Recolectar todos los materiales y OTs
    instalados = {}
    retirados = {}
    luminarias = {}
    ots = set()
    
    # Primero identificar todas las OTs
    for ot in datos_combinados.keys():
        ots.add(ot)
    
    # Ordenar OTs para mantener consistencia
    ots_ordenadas = sorted(ots)
    
    # Ahora procesar todos los datos para cada tipo de material
    for ot, info in datos_combinados.items():
        # Proceso para códigos N1
        for key, nodos_data in info.get('codigos_n1', {}).items():
            nombre = f"LUMINARIA {key}"
            if nombre not in luminarias:
                luminarias[nombre] = {
                    'tipo': 'LUMINARIA',
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            
            # Contar el número total de códigos en todos los nodos para esta OT
            codigos_count = sum(len(codigos) for codigos in nodos_data.values())
            luminarias[nombre]['total'] += codigos_count
            luminarias[nombre]['ots'][ot] += codigos_count
        
        # Proceso para códigos N2
        for key, nodos_data in info.get('codigos_n2', {}).items():
            nombre = f"LUMINARIA {key}"
            if nombre not in luminarias:
                luminarias[nombre] = {
                    'tipo': 'LUMINARIA',
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            
            # Contar el número total de códigos en todos los nodos para esta OT
            codigos_count = sum(len(codigos) for codigos in nodos_data.values())
            luminarias[nombre]['total'] += codigos_count
            luminarias[nombre]['ots'][ot] += codigos_count
            
        # Procesar materiales instalados
        for mat_key, nodo_quantities in info['materiales'].items():
            _, nombre = mat_key.split('|', 1)
            total = sum(nodo_quantities.values())
            
            if nombre not in instalados:
                instalados[nombre] = {
                    'tipo': 'INSTALADO',
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            instalados[nombre]['total'] += total
            instalados[nombre]['ots'][ot] += total
        
        # Procesar materiales retirados
        for mat_key, nodo_quantities in info.get('materiales_retirados', {}).items():
            _, nombre = mat_key.split('|', 1)
            total = sum(nodo_quantities.values())
            
            if nombre not in retirados:
                retirados[nombre] = {
                    'tipo': 'RETIRADO',
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            retirados[nombre]['total'] += total
            retirados[nombre]['ots'][ot] += total

    # Verificación de totales y depuración
    for nombre, info in luminarias.items():
        # Recalcular el total sumando todas las cantidades por OT para confirmar
        total_from_ots = sum(info['ots'].values())
        if info['total'] != total_from_ots:
            # Ajustar si hay discrepancia
            info['total'] = total_from_ots
    
    # Crear DataFrames para cada sección
    columns = ['Nombre Material', 'Unidad', 'Cantidad Total'] + [f'OT_{ot}' for ot in ots_ordenadas]
    
    # Si no hay datos, salir
    if not (luminarias or instalados or retirados):
        return

    # Escribir directamente al Excel sin usar pandas para mayor control
    sheet_name = 'Resumen_general'
    worksheet = writer.book.create_sheet(sheet_name)
    
    # Configuración de estilos
    header_font = Font(bold=True)
    section_font = Font(bold=True, size=12)
    section_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # Escribir encabezados de columnas
    for col, header in enumerate(columns, 1):
        cell = worksheet.cell(row=1, column=col, value=header)
        cell.font = header_font
    
    current_row = 2
    
    # Función para escribir una sección
    def write_section(title, data, row):
        # Encabezado de sección
        cell = worksheet.cell(row=row, column=1, value=title)
        cell.font = section_font
        cell.fill = section_fill
        cell.alignment = center_alignment
        
        # Combinar celdas para el encabezado
        worksheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(columns))
        
        row += 1
        
        # Escribir datos
        for nombre, info in sorted(data.items()):
            worksheet.cell(row=row, column=1, value=nombre)
            worksheet.cell(row=row, column=2, value=info['unidad'])
            worksheet.cell(row=row, column=3, value=info['total'])
            
            # Cantidades por OT
            for i, ot in enumerate(ots_ordenadas, 4):
                worksheet.cell(row=row, column=i, value=info['ots'].get(ot, 0))
            
            row += 1
            
        return row

    # Escribir sección de luminarias
    if luminarias:
        current_row = write_section("LUMINARIAS INSTALADAS", luminarias, current_row)
        current_row += 1  # Espacio entre secciones
    
    # Escribir sección de materiales instalados
    if instalados:
        current_row = write_section("MATERIALES INSTALADOS", instalados, current_row)
        current_row += 1  # Espacio entre secciones
    
    # Escribir sección de materiales retirados
    if retirados:
        current_row = write_section("MATERIALES RETIRADOS", retirados, current_row)
    
    # Ajustar anchos de columna
    for idx, col in enumerate(columns, 1):
        col_width = len(str(col)) + 2
        # Para la primera columna, usar un ancho mayor
        if idx == 1:
            col_width = 40
        worksheet.column_dimensions[get_column_letter(idx)].width = col_width
    
    # Configurar congelación de paneles
    worksheet.freeze_panes = 'D2'
        
def generate_barrio_sheet(writer, barrio_data, barrio_name):
    
    # Normalizar nombre de hoja
    sheet_name = f"{barrio_name.strip().replace('/', '_')[:25]}_R"[:31]
    
    # Recolectar datos del barrio
    materials = {}
    ots = set()
    
    # Procesar materiales instalados
    for mat_key, ot_quantities in barrio_data['materiales_instalados'].items():
        _, nombre = mat_key.split('|', 1)
        for ot, cantidad in ot_quantities.items():
            ots.add(ot)
            key = (nombre, 'INSTALADO')
            
            if key not in materials:
                materials[key] = {
                    'nombre': nombre,
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            materials[key]['total'] += cantidad
            materials[key]['ots'][ot] += cantidad
    
    # Procesar materiales retirados
    for mat_key, ot_quantities in barrio_data['materiales_retirados'].items():
        _, nombre = mat_key.split('|', 1)
        for ot, cantidad in ot_quantities.items():
            ots.add(ot)
            key = (f"{nombre} (RETIRADO)", 'RETIRADO')
            
            if key not in materials:
                materials[key] = {
                    'nombre': key[0],
                    'unidad': 'UND',
                    'total': 0,
                    'ots': defaultdict(int)
                }
            materials[key]['total'] += cantidad
            materials[key]['ots'][ot] += cantidad

    # Crear DataFrame
    rows = []
    for key, data in materials.items():
        row = {
            'Nombre Material': data['nombre'],
            'Unidad': data['unidad'],
            'Cantidad Total': data['total']
        }
        
        # Agregar columnas por OT
        for ot in ots:
            row[f'OT_{ot}'] = data['ots'].get(ot, 0)
        
        rows.append(row)
    
    if not rows:
        return

    df = pd.DataFrame(rows)
    columns = ['Nombre Material', 'Unidad', 'Cantidad Total'] + sorted([f'OT_{ot}' for ot in ots], key=lambda x: x.split('_')[1])
    df = df[columns]

    # Escribir en Excel
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Formatear hoja
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Ajustar anchos de columna
    for idx, col in enumerate(df.columns, 1):
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(col))
        worksheet.column_dimensions[get_column_letter(idx)].width = max_len + 2

    # Congelar paneles
    worksheet.freeze_panes = 'D2'

def cargar_plantilla_mano_obra():
    """
    Carga la plantilla de mano de obra desde un archivo Excel y la modifica para incluir
    partidas unificadas para luminarias.
    
    Returns:
        list: Lista de diccionarios con las partidas de mano de obra
    """
    try:
        # Cargar desde archivo en el mismo directorio
        plantilla_path = "plantilla_mano_obra.xlsx"
        
        if not os.path.exists(plantilla_path):
            raise FileNotFoundError("Archivo de plantilla no encontrado")
            
        df = pd.read_excel(plantilla_path)
        
        # Validar estructura
        required_columns = [
            'DESCRIPCION MANO DE OBRA',
            'UNIDAD', 
            'CANTIDAD',
        ]
        
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Plantilla no tiene las columnas requeridas")
        
        # Convertir a lista de diccionarios
        plantilla_original = df.to_dict('records')
        
        # Modificar la plantilla para unificar los tipos de instalación de luminarias
        plantilla_modificada = []
        
        # Banderas para controlar la adición de partidas unificadas
        tiene_instalacion_luminarias = False
        tiene_desmontaje_luminarias = False
        tiene_transporte_luminarias = False
        
        # Primero, revisar qué tipos de partidas existen en la plantilla original
        for partida in plantilla_original:
            descripcion = partida['DESCRIPCION MANO DE OBRA'].upper()
            
            # Verificar si hay partidas de instalación de luminarias
            #if any(tipo in descripcion for tipo in [
            #    "INSTALACION DE LUMINARIAS EN CAMIONETA", 
            #    "INSTALACION DE LUMINARIAS EN CANASTA", 
            #    "INSTALACION DE LUMINARIAS HORIZONTAL ADOSADA"
            #]):
            tiene_instalacion_luminarias = True
            
            # Verificar si hay partidas de desmontaje de luminarias
            if "DESMONTAJE" in descripcion and "LUMINARIA" in descripcion:
                tiene_desmontaje_luminarias = True
            
            # Verificar si ya existe una partida de transporte de luminarias
            elif "TRANSPORTE" in descripcion and "LUMINARIA" in descripcion:
                tiene_transporte_luminarias = True
        
        # Ahora, crear la plantilla modificada
        for partida in plantilla_original:
            descripcion = partida['DESCRIPCION MANO DE OBRA'].upper()
            
            # Omitir los tipos específicos de instalación de luminarias
            #if any(tipo in descripcion for tipo in [
            #    "INSTALACION DE LUMINARIAS EN CAMIONETA", 
            #    "INSTALACION DE LUMINARIAS EN CANASTA", 
            #    "INSTALACION DE LUMINARIAS HORIZONTAL ADOSADA"
            #]):
            #    # No agregar estas partidas individualmente
            #    continue
            
            # Omitir los tipos específicos de desmontaje de luminarias
            if "DESMONTAJE" in descripcion and "LUMINARIA" in descripcion and (
                "CAMIONETA" in descripcion or "CANASTA" in descripcion):
                # No agregar estas partidas individualmente
                continue
            
            # Mantener otras partidas sin cambios
            else:
                plantilla_modificada.append(partida)
        
        # Agregar partida unificada de instalación de luminarias
        if tiene_instalacion_luminarias:
            plantilla_modificada.append({
                'DESCRIPCION MANO DE OBRA': "INSTALACION LUMINARIAS ESCALERA/CANASTA",
                'UNIDAD': "UND",
                'CANTIDAD': 0  # Se calculará después según el número de nodos
            })
        
        # Agregar partida unificada de desmontaje de luminarias
        if tiene_desmontaje_luminarias:
            plantilla_modificada.append({
                'DESCRIPCION MANO DE OBRA': "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA",
                'UNIDAD': "UND",
                'CANTIDAD': 0
            })
        
        # Agregar partida de transporte de luminarias si no existe
        if not tiene_transporte_luminarias:
            plantilla_modificada.append({
                'DESCRIPCION MANO DE OBRA': "TRANSPORTE DE LUMINARIAS",
                'UNIDAD': "UND",
                'CANTIDAD': 0  # Se calculará después como 2 por nodo
            })
        
        return plantilla_modificada
        
    except Exception as e:
        logger.error(f"Error cargando plantilla: {str(e)}")
        return []

def plantilla_mano_obra(worksheet, df, plantilla):
    """
    Agrega la plantilla de mano de obra a la hoja de Excel.
    
    Args:
        worksheet: Hoja de Excel activa
        df: DataFrame con los datos
        plantilla: Plantilla de mano de obra
    """
    header_font = Font(bold=True)
    cell_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Determinar posición inicial
    start_row = len(df) + 4  # Usar len(df) en lugar de df.shape[0]
    
    # Escribir encabezados
    headers = ['DESCRIPCION MANO DE OBRA', 'UNIDAD', 'CANTIDAD']
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=start_row, column=col_num)
        cell.value = header
        cell.font = header_font
        cell.border = cell_border
    
    # Escribir datos
    for idx, item in enumerate(plantilla, 1):
        row_num = start_row + idx
        descripcion = item['DESCRIPCION MANO DE OBRA']
        
        # Reemplazar "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA" con "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA"
        if descripcion == "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA":
            descripcion = "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA"
            
        worksheet.cell(row=row_num, column=1, value=descripcion).border = cell_border
        worksheet.cell(row=row_num, column=2, value=item['UNIDAD']).border = cell_border
        worksheet.cell(row=row_num, column=3, value=item['CANTIDAD']).border = cell_border
    
    # Ajustar anchos de columna
    worksheet.column_dimensions['A'].width = 45
    worksheet.column_dimensions['B'].width = 10
    worksheet.column_dimensions['C'].width = 12

def agregar_tabla_mano_obra(sheet, df, plantilla):
    """
    Agrega la tabla de mano de obra a la hoja de Excel.
    
    Args:
        sheet: Hoja de Excel activa
        df: DataFrame con los datos de la OT
        plantilla: Plantilla de mano de obra
    """
    # Obtener la OT de la primera fila
    ot = df.iloc[0, 0]
    
    # Obtener información de materiales y postes
    materiales_instalados = {}
    materiales_retirados = {}
    nodos = []
    
    # Extraer los nombres de las columnas que contienen "Nodo_"
    nodo_cols = [col for col in df.columns if col.startswith("Nodo_")]
    
    # Extraer los nodos de la segunda fila (Nodos postes)
    postes_row = df.iloc[1]
    for i, col in enumerate(nodo_cols):
        if pd.notna(postes_row[col]) and postes_row[col]:
            nodo = f"{postes_row[col]}_{i+1}"  # Formato: poste_índice
            nodos.append(nodo)
    
    # Ordenar los nodos por su identificador numérico (si es posible)
    try:
        # Intentar ordenar por el valor numérico del poste
        nodos = sorted(nodos, key=lambda x: int(x.split('_')[0]))
    except (ValueError, TypeError):
        # Si no se puede ordenar numéricamente, mantener el orden original
        pass
    
    # Procesar materiales instalados
    in_materiales = False
    in_retirados = False
    
    for _, row in df.iterrows():
        if row.iloc[0] == "MATERIALES INSTALADOS":
            in_materiales = True
            in_retirados = False
            continue
        elif row.iloc[0] == "MATERIALES RETIRADOS":
            in_materiales = False
            in_retirados = True
            continue
        elif row.iloc[0] in ["OBSERVACIONES POR NODO", "OBSERVACIONES COMPLETAS"]:
            in_materiales = False
            in_retirados = False
            continue
        
        if in_materiales and pd.notna(row.iloc[0]) and row.iloc[0] != "":
            material_key = f"material|{row.iloc[0]}"
            materiales_instalados[material_key] = {}
            for i, nodo in enumerate(nodos):
                if i < len(nodo_cols):  # Asegurarse de que no exceda el límite de columnas
                    qty = row[nodo_cols[i]]
                    if pd.notna(qty) and qty > 0:
                        materiales_instalados[material_key][nodo] = float(qty)
        
        elif in_retirados and pd.notna(row.iloc[0]) and row.iloc[0] != "":
            material_key = f"material|{row.iloc[0]}"
            materiales_retirados[material_key] = {}
            for i, nodo in enumerate(nodos):
                if i < len(nodo_cols):  # Asegurarse de que no exceda el límite de columnas
                    qty = row[nodo_cols[i]]
                    if pd.notna(qty) and qty > 0:
                        materiales_retirados[material_key][nodo] = float(qty)
    
    # Encontrar la última fila utilizada en la hoja
    last_row = sheet.max_row + 2  # Dejamos una fila de espacio
    
    # Agregar cabecera de la tabla de mano de obra
    sheet.cell(row=last_row, column=1, value="TABLA DE MANO DE OBRA").font = Font(bold=True, size=14)
    sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
    last_row += 2
    
    # Encabezados de la tabla
    headers = ["DESCRIPCIÓN MANO DE OBRA", "UNIDAD", "CANTIDAD", "NODO", "MATERIALES ASOCIADOS"]
    for i, header in enumerate(headers, 1):
        sheet.cell(row=last_row, column=i, value=header).font = Font(bold=True)
        sheet.cell(row=last_row, column=i).fill = PatternFill("solid", fgColor="DDDDDD")
        sheet.cell(row=last_row, column=i).border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
    last_row += 1
    
    # Agrupar partidas por bloques para mejor visualización
    bloques = [
        ("Postes",
         ["APERTURA", "APLOMADA", "CONCRETADA", "HINCADA"],
         ["POSTE"]),
        ("Instalación luminarias",
         ["INSTALACION LUMINARIAS"],
         ["LUMINARIA", "FOTOCELDA", "GRILLETE", "BRAZO"]),
        ("Conexión a tierra",
         ["CONEXIÓN A CABLE A TIERRA", "INSTALACION KIT SPT"],
         ["KIT DE PUESTA A TIERRA", "CONECT PERF", "CONECTOR BIME/COM", "ALAMBRE", "TUERCA", "TORNILLO"]),
        ("Desmontaje / Transporte",
         ["DESMONTAJE", "TRANSPORTE"],
         ["ALAMBRE", "BRAZO", "CÓDIGO", "CABLE"])
    ]
    
    # Pre-procesamiento para unificar los tipos de desmontaje de luminarias
    partidas_unificadas = []
    tipo_desmontaje_luminarias = []
    
    for partida in plantilla:
        descripcion = partida['DESCRIPCION MANO DE OBRA']
        # Verificar si es una partida de desmontaje de luminarias
        if ("DESMONTAJE" in descripcion.upper() and "LUMINARIA" in descripcion.upper() and 
            ("CAMIONETA" in descripcion.upper() or "CANASTA" in descripcion.upper())):
            # Guardar la descripción original para referencia
            tipo_desmontaje_luminarias.append(descripcion)
            # No agregar esta partida a la lista unificada aún
        else:
            partidas_unificadas.append(partida)
    
    # Si hay partidas de desmontaje de luminarias, crear una partida unificada
    if tipo_desmontaje_luminarias:
        partida_unificada = {
            'DESCRIPCION MANO DE OBRA': "DESMONTAJE DE LUMINARIAS UNIFICADO",
            'UNIDAD': next((p['UNIDAD'] for p in plantilla if p['DESCRIPCION MANO DE OBRA'] in tipo_desmontaje_luminarias), "UN")
        }
        partidas_unificadas.append(partida_unificada)
    
    # Reemplazar la plantilla original con la unificada
    plantilla_original = plantilla
    plantilla = partidas_unificadas
    
    # Para cada nodo, mostrar todos los bloques de mano de obra en una mejor organización
    for nodo in nodos:
        # Obtener poste ID para mostrar
        poste_id = nodo.split('_')[0]
        
        # Insertar encabezado de nodo con mejor diseño
        sheet.cell(row=last_row, column=1, value=f"NODO: {poste_id}").font = Font(bold=True, size=12)
        sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
        sheet.cell(row=last_row, column=1).fill = PatternFill("solid", fgColor="3366FF")  # Azul
        sheet.cell(row=last_row, column=1).font = Font(bold=True, color="FFFFFF")  # Texto blanco
        sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
        last_row += 1
        
        # Contador para comprobar si hay alguna partida para este nodo
        partidas_nodo_count = 0
        
        # Para cada bloque de partidas
        for titulo_bloque, keywords_mo, keywords_mat in bloques:
            # Variables para rastrear si ya hemos añadido el título del bloque
            bloque_agregado = False
            
            # Filtrar partidas del bloque actual
            partidas_filtradas = [
                item for item in plantilla
                if any(kw in item['DESCRIPCION MANO DE OBRA'].upper() for kw in keywords_mo)
            ]
            
            # Para cada partida, calcular la mano de obra necesaria para este nodo
            for partida in partidas_filtradas:
                descripcion = partida['DESCRIPCION MANO DE OBRA']
                unidad = partida['UNIDAD']
                
                # Determinar si se trata de una partida específica para cierto tipo de poste
                tipo_poste = None
                if "METALICO" in descripcion.upper() or "METÁLICO" in descripcion.upper():
                    tipo_poste = "METALICO"
                elif "FIBRA" in descripcion.upper():
                    tipo_poste = "FIBRA"
                elif "CONCRETO" in descripcion.upper() or "HORMIGÓN" in descripcion.upper():
                    tipo_poste = "CONCRETO"
                
                # Caso especial para desmontaje de luminarias unificado
                cantidad_mo = 0
                materiales_inst = []
                materiales_ret = []
                
                if descripcion == "DESMONTAJE DE LUMINARIAS UNIFICADO":
                    # Calcular la cantidad sumando todos los tipos de desmontaje
                    for desc_original in tipo_desmontaje_luminarias:
                        cant_temp, mat_inst_temp, mat_ret_temp = calcular_cantidad_mano_obra(
                            desc_original,
                            materiales_instalados,
                            materiales_retirados,
                            nodo
                        )
                        cantidad_mo += cant_temp
                        materiales_inst.extend(mat_inst_temp)
                        materiales_ret.extend(mat_ret_temp)
                    
                    # Eliminar duplicados en las listas de materiales
                    materiales_inst = list(set(materiales_inst))
                    materiales_ret = list(set(materiales_ret))
                else:
                    # Procesamiento normal para otras partidas
                    cantidad_mo, materiales_inst, materiales_ret = calcular_cantidad_mano_obra(
                        descripcion,
                        materiales_instalados,
                        materiales_retirados,
                        nodo
                    )
                
                # Solo agregar filas para partidas con cantidad > 0 
                if cantidad_mo > 0:
                    # Agregar título del bloque si aún no se ha hecho y hay partidas
                    if not bloque_agregado:
                        sheet.cell(row=last_row, column=1, value=titulo_bloque).font = Font(bold=True)
                        sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
                        sheet.cell(row=last_row, column=1).fill = PatternFill("solid", fgColor="AAAAAA")
                        sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
                        last_row += 1
                        bloque_agregado = True
                    
                    partidas_nodo_count += 1
                    
                    # Descripción de la mano de obra
                    sheet.cell(row=last_row, column=1, value=descripcion).border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    
                    # Unidad
                    sheet.cell(row=last_row, column=2, value=unidad).border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    sheet.cell(row=last_row, column=2).alignment = Alignment(horizontal='center')
                    
                    # Cantidad
                    sheet.cell(row=last_row, column=3, value=cantidad_mo).border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    sheet.cell(row=last_row, column=3).alignment = Alignment(horizontal='center')
                    
                    # Nodo
                    # Extraer el número de poste del nodo
                    poste = nodo.split('_')[0]
                    sheet.cell(row=last_row, column=4, value=poste).border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    sheet.cell(row=last_row, column=4).alignment = Alignment(horizontal='center')
                    
                    # Materiales asociados
                    materiales_texto = []
                    if materiales_inst:
                        materiales_texto.append("INST: " + ", ".join(materiales_inst))
                    if materiales_ret:
                        materiales_texto.append("RET: " + ", ".join(materiales_ret))
                    
                    sheet.cell(row=last_row, column=5, value="\n".join(materiales_texto)).border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                    sheet.cell(row=last_row, column=5).alignment = Alignment(vertical='top', wrap_text=True)
                    
                    last_row += 1
            
            # Solo añadir espacio si se agregó alguna partida en este bloque
            if bloque_agregado:
                last_row += 1
        
        # Si no hay partidas para este nodo, mostrar mensaje indicativo
        if partidas_nodo_count == 0:
            sheet.cell(row=last_row, column=1, value="No hay partidas de mano de obra asociadas a este nodo").border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
            sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
            last_row += 1
        
        # Espacio después de cada nodo para mejor visualización
        last_row += 2
    
    # Resumen general de mano de obra (agregar después de mostrar por nodos)
    sheet.cell(row=last_row, column=1, value="RESUMEN GENERAL DE MANO DE OBRA").font = Font(bold=True, size=14)
    sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
    sheet.cell(row=last_row, column=1).fill = PatternFill("solid", fgColor="009900")  # Verde
    sheet.cell(row=last_row, column=1).font = Font(bold=True, color="FFFFFF")  # Texto blanco
    last_row += 2
    
    # Encabezados para el resumen
    for i, header in enumerate(["DESCRIPCIÓN MANO DE OBRA", "UNIDAD", "CANTIDAD TOTAL", "", ""], 1):
        sheet.cell(row=last_row, column=i, value=header).font = Font(bold=True)
        sheet.cell(row=last_row, column=i).fill = PatternFill("solid", fgColor="DDDDDD")
        sheet.cell(row=last_row, column=i).border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
    last_row += 1
    
    # Diccionario para acumular totales
    totales_mo = {}
    
    # Acumular totales por descripción de mano de obra
    for nodo in nodos:
        # Procesamiento especial para desmontaje de luminarias unificado
        if any(("DESMONTAJE" in partida['DESCRIPCION MANO DE OBRA'].upper() and 
                "LUMINARIA" in partida['DESCRIPCION MANO DE OBRA'].upper() and 
                ("CAMIONETA" in partida['DESCRIPCION MANO DE OBRA'].upper() or 
                 "CANASTA" in partida['DESCRIPCION MANO DE OBRA'].upper())) 
               for partida in plantilla_original):
            
            desc_unificada = "DESMONTAJE DE LUMINARIAS UNIFICADO"
            unidad = next((p['UNIDAD'] for p in plantilla if p['DESCRIPCION MANO DE OBRA'] == desc_unificada), "UN")
            cantidad_total = 0
            
            # Sumar cantidades de todos los tipos de desmontaje
            for desc_original in tipo_desmontaje_luminarias:
                cant_temp, _, _ = calcular_cantidad_mano_obra(
                    desc_original,
                    materiales_instalados,
                    materiales_retirados,
                    nodo
                )
                cantidad_total += cant_temp
                
            if cantidad_total > 0:
                if (desc_unificada, unidad) not in totales_mo:
                    totales_mo[(desc_unificada, unidad)] = 0
                totales_mo[(desc_unificada, unidad)] += cantidad_total
        
        # Procesamiento normal para otras partidas
        for partida in plantilla:
            descripcion = partida['DESCRIPCION MANO DE OBRA']
            unidad = partida['UNIDAD']
            
            # Saltar el procesamiento del desmontaje unificado que ya se procesó arriba
            if descripcion == "DESMONTAJE DE LUMINARIAS UNIFICADO":
                continue
                
            cantidad_mo, _, _ = calcular_cantidad_mano_obra(
                descripcion,
                materiales_instalados,
                materiales_retirados,
                nodo
            )
            
            if cantidad_mo > 0:
                if (descripcion, unidad) not in totales_mo:
                    totales_mo[(descripcion, unidad)] = 0
                totales_mo[(descripcion, unidad)] += cantidad_mo
    
    # Mostrar totales agrupados por tipos de partidas
    for titulo_bloque, keywords_mo, _ in bloques:
        # Insertar título del bloque
        sheet.cell(row=last_row, column=1, value=titulo_bloque).font = Font(bold=True)
        sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
        sheet.cell(row=last_row, column=1).fill = PatternFill("solid", fgColor="AAAAAA")
        sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
        last_row += 1
        
        # Filtrar partidas que pertenecen a este bloque y que tienen totales
        partidas_bloque = [(desc, und, total) for (desc, und), total in totales_mo.items() 
                          if any(kw in desc.upper() for kw in keywords_mo)]
        
        # Ordenar por descripción
        partidas_bloque.sort(key=lambda x: x[0])
        
        # Mostrar partidas
        for desc, und, total in partidas_bloque:
            # Descripción
            sheet.cell(row=last_row, column=1, value=desc).border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            
            # Unidad
            sheet.cell(row=last_row, column=2, value=und).border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            sheet.cell(row=last_row, column=2).alignment = Alignment(horizontal='center')
            
            # Cantidad total
            sheet.cell(row=last_row, column=3, value=total).border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            sheet.cell(row=last_row, column=3).alignment = Alignment(horizontal='center')
            
            last_row += 1
        
        # Espacio después de cada bloque en el resumen
        last_row += 1
    
    # Ajustar ancho de columnas para mejor visualización
    for col in range(1, 6):
        col_letter = get_column_letter(col)
        max_length = 0
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
        
        # Ajustar el ancho con un pequeño margen adicional
        adjusted_width = max_length + 2
        sheet.column_dimensions[col_letter].width = min(adjusted_width, 50)  # Máximo 50 para evitar columnas demasiado anchas

    # Agregar pie de página
    last_row += 2
    #sheet.cell(row=last_row, column=1, value=f"OT: {ot} - Generado automáticamente").font = Font(italic=True)
    sheet.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=5)
    sheet.cell(row=last_row, column=1).alignment = Alignment(horizontal='center')
   
def calcular_cantidad_mano_obra(descripcion, materiales_instalados, materiales_retirados, nodo, codigos_n1=None, codigos_n2=None):
    """
    Calcula la cantidad de mano de obra necesaria para una partida específica en un nodo.
    
    Args:
        descripcion: Descripción de la partida de mano de obra
        materiales_instalados: Diccionario de materiales instalados por nodo
        materiales_retirados: Diccionario de materiales retirados por nodo
        nodo: Nodo actual
        codigos_n1: Diccionario de códigos N1 instalados (opcional)
        codigos_n2: Diccionario de códigos N2 instalados (opcional)
    
    Returns:
        cantidad_mo: Cantidad de mano de obra necesaria
        materiales_inst: Lista de materiales instalados asociados
        materiales_ret: Lista de materiales retirados asociados
    """
    cantidad_mo = 0
    materiales_inst = []
    materiales_ret = []
    
    descripcion_upper = descripcion.upper()
    
    # Unificar todos los tipos de instalación de luminarias bajo un solo tipo
    if "INSTALACION DE CANASTA/ESCALERA" in descripcion_upper or any(tipo in descripcion_upper for tipo in [
        "INSTALACION DE LUMINARIAS EN CAMIONETA", 
        "INSTALACION DE LUMINARIAS EN CANASTA", 
        "INSTALACION DE LUMINARIAS HORIZONTAL ADOSADA",
        "INSTALACION LUMINARIAS"
    ]):
        # Lista ampliada de palabras clave para detectar luminarias - similar a la lógica de transporte
        luminaria_keywords = [
            "LUMINARIA", "LAMPARA", "LED", "BOMBILLA", "PROYECTOR", 
            "REFLECTOR", "FOCO", "BALASTRO", "CODIGO 1", "CODIGO 2",
            "CODIGO N1", "CODIGO N2", "CODIGO DE LUMINARIA"
        ]
        
        # Inicializar contador de luminarias instaladas
        luminarias_instaladas_count = 0
        
        # 1. Verificar códigos N1 y N2 instalados
        if codigos_n1:
            for key_codigo, nodos_valores in codigos_n1.items():
                if nodo in nodos_valores:
                    cantidad_codigos = len(nodos_valores[nodo])
                    luminarias_instaladas_count += cantidad_codigos
                    if cantidad_codigos > 0:
                        materiales_inst.append(f"CÓDIGO N1: {', '.join(nodos_valores[nodo])}")
        
        if codigos_n2:
            for key_codigo, nodos_valores in codigos_n2.items():
                if nodo in nodos_valores:
                    cantidad_codigos = len(nodos_valores[nodo])
                    luminarias_instaladas_count += cantidad_codigos
                    if cantidad_codigos > 0:
                        materiales_inst.append(f"CÓDIGO N2: {', '.join(nodos_valores[nodo])}")
        
        # 2. Buscar luminarias instaladas en materiales
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if any(kw in material_name for kw in luminaria_keywords) and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    luminarias_instaladas_count += qty
                    materiales_inst.append(f"{material_name} ({qty})")
        
        # La cantidad de mano de obra es igual al número de luminarias instaladas
        cantidad_mo = luminarias_instaladas_count
        
        # Buscar materiales complementarios independientemente de si encontramos luminarias
        complementos = ["FOTOCELDA", "GRILLETE", "BRAZO"]
        for comp_material_key, comp_nodos_qty in materiales_instalados.items():
            if "|" in comp_material_key:
                comp_material_name = comp_material_key.split("|")[1].upper()
                if any(comp in comp_material_name for comp in complementos) and nodo in comp_nodos_qty:
                    materiales_inst.append(comp_material_name)
        
        return cantidad_mo, materiales_inst, materiales_ret
    
    # Unificar todos los tipos de desmontaje de luminarias
    if "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA" in descripcion_upper or (
        "DESMONTAJE" in descripcion_upper and "LUMINARIA" in descripcion_upper and (
        "CAMIONETA" in descripcion_upper or "CANASTA" in descripcion_upper)):
        descripcion_upper = "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA"
    
    # PROCESAMIENTO UNIFICADO PARA TRANSPORTE DE LUMINARIAS
    # Detecta cualquier variante de "transporte" + "luminaria"
    if any(transp in descripcion_upper for transp in ["TRANSPORTE", "TRANSP."]) and any(lum in descripcion_upper for lum in ["LUMINARIA", "LUMINARIAS"]):
        # Inicializar contadores
        luminarias_instaladas_count = 0
        luminarias_retiradas_count = 0
        
        # Lista ampliada de palabras clave para detectar luminarias
        luminaria_keywords = [
            "LUMINARIA", "LAMPARA", "LED", "BOMBILLA", "PROYECTOR", 
            "REFLECTOR", "FOCO", "BALASTRO", "CODIGO 1", "CODIGO 2",
            "CODIGO N1", "CODIGO N2", "CODIGO DE LUMINARIA"
        ]
        
        # 1. Verificar códigos N1 y N2 instalados
        if codigos_n1:
            for key_codigo, nodos_valores in codigos_n1.items():
                if nodo in nodos_valores:
                    cantidad_codigos = len(nodos_valores[nodo])
                    luminarias_instaladas_count += cantidad_codigos
                    if cantidad_codigos > 0:
                        materiales_inst.append(f"CÓDIGO N1: {', '.join(nodos_valores[nodo])}")
        
        if codigos_n2:
            for key_codigo, nodos_valores in codigos_n2.items():
                if nodo in nodos_valores:
                    cantidad_codigos = len(nodos_valores[nodo])
                    luminarias_instaladas_count += cantidad_codigos
                    if cantidad_codigos > 0:
                        materiales_inst.append(f"CÓDIGO N2: {', '.join(nodos_valores[nodo])}")
        
        # 2. Buscar luminarias instaladas en materiales
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if any(kw in material_name for kw in luminaria_keywords) and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    luminarias_instaladas_count += qty
                    materiales_inst.append(f"{material_name} ({qty})")
        
        # 3. Buscar luminarias retiradas - MEJORADO
        for material_key, nodos_qty in materiales_retirados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                # Verificar si es una luminaria retirada usando múltiples criterios
                es_luminaria_retirada = (
                    any(kw in material_name for kw in luminaria_keywords) or 
                    "RETIRADA" in material_name or
                    "LUMINARIA RETIRADA" in material_name or
                    "BOMBILLA RETIRADA" in material_name or
                    "FOTOCELDA RETIRADA" in material_name
                )
                
                if es_luminaria_retirada and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    luminarias_retiradas_count += qty
                    materiales_ret.append(f"{material_name} ({qty})")
        
        # Sumar ambos conteos para transporte total
        cantidad_mo = luminarias_instaladas_count + luminarias_retiradas_count
        
        return cantidad_mo, materiales_inst, materiales_ret
    
    # 5. Mano de obra para transporte (MODIFICADO PARA INCLUIR DIFERENTES ELEMENTOS)
    elif "TRANSPORTE" in descripcion_upper or "TRANSP." in descripcion_upper:
        # Determinar qué tipo de elemento se está transportando
        if "POSTE" in descripcion_upper:
            # Contar postes instalados y retirados
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "POSTE" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
            
            for material_key, nodos_qty in materiales_retirados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "POSTE" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_ret.append(f"{material_name} ({qty})")
        
        elif "BRAZO" in descripcion_upper:
            # Contar brazos instalados y retirados
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "BRAZO" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
            
            for material_key, nodos_qty in materiales_retirados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "BRAZO" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_ret.append(f"{material_name} ({qty})")
        
        elif "CABLE" in descripcion_upper:
            # Sumar metros de cable instalado y retirado
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if any(kw in material_name for kw in ["CABLE", "ALAMBRE"]) and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
            
            for material_key, nodos_qty in materiales_retirados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if any(kw in material_name for kw in ["CABLE", "ALAMBRE"]) and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_ret.append(f"{material_name} ({qty})")
        
        elif "VARILLA" in descripcion_upper or "KIT TIERRA" in descripcion_upper:
            # Contar kits de tierra y varillas
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if any(kw in material_name for kw in ["VARILLA", "KIT DE PUESTA A TIERRA"]) and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
        
        elif "PERCHA" in descripcion_upper:
            # Contar perchas instaladas
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "PERCHA" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
        
        elif "COLLARINES" in descripcion_upper:
            # Contar bandas/collarines instalados
            for material_key, nodos_qty in materiales_instalados.items():
                if "|" in material_key:
                    material_name = material_key.split("|")[1].upper()
                    if "BANDA" in material_name and nodo in nodos_qty:
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(f"{material_name} ({qty})")
    
    # Mano de obra para postes de concreto de 8 a 10 mts
    elif descripcion_upper in [
        "APERTURA HUECOS POSTES ANCLAS SECUNDARIAS DE 8 A 10 MTS",
        "APLOMADA POSTES DE CONCRETO DE 8 A 10 MTS",
        "CONCRETADA DE POSTE CONCRETO DE 8 A 12 M INCLUYE MATERIALES Y MO",
        "HINCADA DE POSTES CONCRETO DE 8 A 12 MTS",
        "TRANSP.POSTE.CONC.10MT.SITIO SIN INCREME"
    ]:
        # Buscar postes de concreto de 8-10 mts instalados
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and "CONCRETO" in material_name and nodo in nodos_qty:
                    # Verificar que sea un poste de 8-10 mts
                    if any(altura in material_name for altura in ["8M", "9M", "10M"]):
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(material_name)
        return cantidad_mo, materiales_inst, materiales_ret
    
    # Mano de obra para postes de concreto de 11 a 14 mts
    elif descripcion_upper in [
        "APERTURA HUECOS POSTES ANCLAS PRIMARIA DE 11 A 14",
        "APLOMADA POSTES DE CONCRETO DE 11 A 14 MTS",
        "CONCRETADA DE POSTE PRIMARIOS",
        "HINCADA DE POSTES DE 14 MTS",
        "TRANSP.POSTE.CONC.12MT.SITIO SIN INCREME"
    ]:
        # Buscar postes de concreto de 11-14 mts instalados
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and "CONCRETO" in material_name and nodo in nodos_qty:
                    # Verificar que sea un poste de 11-14 mts
                    if any(altura in material_name for altura in ["11M", "12M", "13M", "14M"]):
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(material_name)
        return cantidad_mo, materiales_inst, materiales_ret
    
    # Mano de obra para postes metálicos
    elif descripcion_upper in [
        "APERTURA HUECOS POSTES ANCLAS SECUNDARIAS DE 8 A 10 MTS",
        "APLOMADA POSTES METALICOS Y/O FIBRA VIDRIO 8 A 10",
        "APLOMADA POSTES METALICOS Y/O FIBRA VIDRIO 11 A 14",
        "CONCRETADA DE POSTE METALICO 8 A 12 MT INCLUYE MATERIALES Y MO",
        "HINCADA DE POSTE METALICO DE 4 A 8M",
        "HINCADA DE POSTE METALICO DE 10 A 12M",
        "TRANSP.POSTE.METALICO DE 4 A 12MT"
    ]:
        # Determinar si se trata de un poste metálico y su altura
        altura_relacionada = None
        if "4 A 8M" in descripcion_upper:
            altura_relacionada = ["4M", "5M", "6M", "7M", "8M"]
        elif "8 A 10" in descripcion_upper or "10 A 12M" in descripcion_upper:
            altura_relacionada = ["8M", "9M", "10M", "11M", "12M"]
        elif "11 A 14" in descripcion_upper:
            altura_relacionada = ["11M", "12M", "13M", "14M"]
        
        # Buscar postes metálicos instalados con la altura correspondiente
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and "METALICO" in material_name and nodo in nodos_qty:
                    # Si se especificó altura y coincide o si no importa la altura
                    if not altura_relacionada or any(altura in material_name for altura in altura_relacionada):
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(material_name)
        return cantidad_mo, materiales_inst, materiales_ret
    
    # Mano de obra para postes de fibra
    elif descripcion_upper in [
        "APERTURA HUECOS POSTES ANCLAS SECUNDARIAS DE 8 A 10 MTS",
        "APLOMADA POSTES METALICOS Y/O FIBRA VIDRIO 8 A 10",
        "APLOMADA POSTES METALICOS Y/O FIBRA VIDRIO 11 A 14",
        "CONCRETADA DE POSTE FIBRA 8 A 12 MT INCLUYE MATERIALES",
        "HINCADA DE POSTE FIBRA DE 8M",
        "HINCADA DE POSTE FIBRA DE 10 A 12M",
        "TRANSP.POSTE.METALICO DE 4 A 12MT"  # Se usa el mismo transporte
    ]:
        # Determinar si se trata de un poste de fibra y su altura
        altura_relacionada = None
        if "8M" in descripcion_upper:
            altura_relacionada = ["8M"]
        elif "10 A 12M" in descripcion_upper:
            altura_relacionada = ["10M", "11M", "12M"]
        
        # Buscar postes de fibra instalados con la altura correspondiente
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and "FIBRA" in material_name and nodo in nodos_qty:
                    # Si se especificó altura y coincide o si no importa la altura
                    if not altura_relacionada or any(altura in material_name for altura in altura_relacionada):
                        qty = nodos_qty[nodo]
                        cantidad_mo += qty
                        materiales_inst.append(material_name)
        return cantidad_mo, materiales_inst, materiales_ret
    
    # Alquiler de Machine Compresor (para todos los postes)
    elif descripcion_upper == 'ALQUILER DE EQUIPO "MACHINE COMPRESOR"':
        # Contar todos los postes instalados (de cualquier tipo)
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    cantidad_mo += qty
                    materiales_inst.append(material_name)
        return cantidad_mo, materiales_inst, materiales_ret
    
    # 1. Mano de obra relacionada con postes (casos genéricos no específicos)
    if any(keyword in descripcion_upper for keyword in ["APERTURA", "APLOMADA", "CONCRETADA", "HINCADA"]):
        # Determinar el tipo de poste según la descripción
        tipo_poste_desc = None
        altura_poste = None
        
        if "METALICO" in descripcion_upper or "METÁLICO" in descripcion_upper:
            tipo_poste_desc = "METALICO"
        elif "FIBRA" in descripcion_upper:
            tipo_poste_desc = "FIBRA"
        elif "CONCRETO" in descripcion_upper or "HORMIGÓN" in descripcion_upper:
            tipo_poste_desc = "CONCRETO"
            
        if "8 A 10" in descripcion_upper or "8 A 12" in descripcion_upper:
            altura_poste = ["8M", "9M", "10M", "11M", "12M"]
        elif "11 A 14" in descripcion_upper:
            altura_poste = ["11M", "12M", "13M", "14M"]
        elif "4 A 8" in descripcion_upper:
            altura_poste = ["4M", "5M", "6M", "7M", "8M"]
            
        # Buscar materiales relacionados con postes
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "POSTE" in material_name and nodo in nodos_qty:
                    # Verificar coincidencia de tipo si está especificado
                    if tipo_poste_desc and tipo_poste_desc not in material_name:
                        continue
                        
                    # Verificar coincidencia de altura si está especificada
                    if altura_poste and not any(alt in material_name for alt in altura_poste):
                        continue
                    
                    qty = nodos_qty[nodo]
                    if qty > 0:
                        cantidad_mo += qty
                        materiales_inst.append(material_name)
    
    # 3. Mano de obra para conexión a tierra e instalación de kit SPT
    elif any(keyword in descripcion_upper for keyword in ["CONEXIÓN A CABLE A TIERRA", "INSTALACION DE ATERRIZAJES"]):
        # Verificar primero si hay un kit de puesta a tierra instalado en este nodo
        tiene_kit_spt = False
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if "KIT DE PUESTA A TIERRA" in material_name and nodo in nodos_qty and nodos_qty[nodo] > 0:
                    tiene_kit_spt = True
                    break
        
        # Si hay un kit SPT instalado, no cobrar la conexión a tierra
        if tiene_kit_spt:
            return 0, [], []
        
        # Si no hay kit SPT, proceder con la conexión a tierra normal
        kit_keywords = ["CONECT PERF", "CONECTOR BIME", "ALAMBRE", "VARILLA", "TIERRA", "CONECTOR VARILLA"]
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if any(kw in material_name for kw in kit_keywords) and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    if "VARILLA" in material_name:  # Solo contar la varilla principal una vez
                        cantidad_mo += qty
                    materiales_inst.append(material_name)
    
    # 3.1 Mano de obra para instalación de kit SPT
    elif "INSTALACION KIT SPT" in descripcion_upper:
        kit_keywords = ["KIT DE PUESTA A TIERRA", "VARILLA"]
        for material_key, nodos_qty in materiales_instalados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if any(kw in material_name for kw in kit_keywords) and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    cantidad_mo += qty
                    materiales_inst.append(material_name)
    
    # 4. Mano de obra para desmontaje (diferente al unificado)
    elif "DESMONTAJE" in descripcion_upper:
        # Determinar qué tipo de elemento se está desmontando
        keywords = []
        if "LUMINARIA" in descripcion_upper:
            keywords = ["LUMINARIA", "BOMBILLA", "LED"]
        elif "POSTE" in descripcion_upper:
            keywords = ["POSTE"]
        elif "ALAMBRE" in descripcion_upper or "CABLE" in descripcion_upper:
            keywords = ["ALAMBRE", "CABLE", "CONDUCTOR"]
        elif "BRAZO" in descripcion_upper:
            keywords = ["BRAZO"]
        
        for material_key, nodos_qty in materiales_retirados.items():
            if "|" in material_key:
                material_name = material_key.split("|")[1].upper()
                if any(kw in material_name for kw in keywords) and nodo in nodos_qty:
                    qty = nodos_qty[nodo]
                    cantidad_mo += qty
                    materiales_ret.append(material_name)
    
    # 6. Instalación de cable secundario
    elif "INSTALACION CABLE SECUNDARIO" in descripcion_upper:
        # Determinar si es aéreo o subterráneo
        tipo_instalacion = None
        if "AEREO" in descripcion_upper:
            tipo_instalacion = "AEREO"
        elif "SUBTERRANEO" in descripcion_upper:
            tipo_instalacion = "SUBTERRANEO"
            
        # Buscar cables instalados
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["CABLE", "TPX", "AL #"]) and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    # 7. Desmontaje de cable secundario
    elif "DESMONTAJE CABLE SECUNDARIO" in descripcion_upper:
        # Buscar cables retirados
        for material_key, nodos_qty in materiales_retirados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["CABLE", "RETIRADO"]) and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_ret.append(material_name)
    
    # 8. Instalación de trama de seguridad
    elif "INSTALACION TRAMA DE SEGURIDAD" in descripcion_upper:
        # Contar luminarias instaladas
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["LUMINARIA", "LED", "CODIGO LUMINARIA"]) and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    # 9. Vestida conjunto perchas
    elif "VESTIDA CONJUNTO" in descripcion_upper:
        # Contar perchas instaladas
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if "PERCHA" in material_name and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    # 10. Instalación de cajas de AP
    elif "CAJA DE A.P" in descripcion_upper:
        # Contar cajas instaladas
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if "CAJA DE A.P" in material_name and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    # 11. Pintada de nodos
    elif "PINTADA DE NODOS" in descripcion_upper:
        # Contar postes instalados
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if "POSTE" in material_name and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    # 12. Excavación y recuperación de zona
    elif "EXCAVACION" in descripcion_upper or "RECUPERACION ZONA" in descripcion_upper:
        # Para estos casos, la cantidad se determina por otros medios
        # y generalmente se especifica directamente en la tabla
        # Aquí podríamos asociar materiales como tubos o cables subterráneos
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["TUBO", "TUBERIA", "CONDUIT"]) and nodo in nodos_qty:
                materiales_inst.append(material_name)
    
    # 13. Soldadura por punto
    elif "SOLDADURA" in descripcion_upper:
        # Para estos casos, la cantidad se determina por otros medios
        # Podríamos asociar materiales como conectores o empalmes
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["CONECTOR", "EMPALME", "GEL"]) and nodo in nodos_qty:
                materiales_inst.append(material_name)
    
    # 14. Instalación de coraza o conduflex
    elif "INSTALACION CORAZA" in descripcion_upper:
        # Buscar materiales de coraza o conduflex
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["CORAZA", "CONDUFLEX"]) and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    # 15. Instalación de empalmes o conectores
    elif "INSTALACION DE CUBIERTA GEL" in descripcion_upper or "INSTALACION DE EMPALME" in descripcion_upper:
        # Buscar empalmes o conectores
        for material_key, nodos_qty in materiales_instalados.items():
            material_name = material_key.split("|")[1].upper()
            if any(kw in material_name for kw in ["EMPALME GEL", "CONECTOR", "GEL"]) and nodo in nodos_qty:
                qty = nodos_qty[nodo]
                cantidad_mo += qty
                materiales_inst.append(material_name)
    
    return cantidad_mo, materiales_inst, materiales_ret

def extraer_cantidad(texto):
    """
    Extrae la cantidad numérica de un texto.
    
    Formatos soportados:
    - 'Descripción (cantidad)'
    - 'Descripción cantidad UND'
    - 'cantidad UND'
    
    Args:
        texto: Texto del que extraer la cantidad
        
    Returns:
        float: Cantidad extraída o 0 si no se encuentra
    """
    try:
        # Caso 1: Formato 'Descripción (cantidad)'
        if '(' in texto and ')' in texto:
            cantidad_str = texto.split('(')[-1].rstrip(')')
            return float(cantidad_str)
        
        # Caso 2: Buscar patrón de número seguido de UND
        palabras = texto.split()
        for i, palabra in enumerate(palabras):
            if palabra.upper() in ["UND", "UN", "ML", "M", "KG", "MT", "MTS"]:
                if i > 0 and palabras[i-1].replace(',', '.').replace('-', '').isdigit():
                    return float(palabras[i-1].replace(',', '.'))
                elif i > 0:
                    try:
                        return float(palabras[i-1].replace(',', '.'))
                    except ValueError:
                        pass
        
        # Caso 3: Intentar encontrar cualquier número en el texto
        import re
        numeros = re.findall(r'\d+(?:\.\d+)?', texto.replace(',', '.'))
        if numeros:
            return float(numeros[0])
            
        return 0
    except ValueError:
        return 0  # En caso de error, devolver 0

def generar_excel(datos_combinados, datos_por_barrio_combinados, dfs_originales_combinados):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        temp_sheet_name = "PlantillaInicial"
        writer.book.create_sheet(temp_sheet_name)
        sheets_created = False        
        
        generate_resumen_general(writer, datos_combinados)
        
        generate_resumen_tecnicos(writer, datos_combinados, dfs_originales_combinados)  
        
        #agregar_hoja_asociaciones(writer, datos_combinados) 
                                 
        if datos_combinados:                        
            
            for ot, info in datos_combinados.items():
                seen = set()
                nodos_con_fechas = [(nodo, pd.to_datetime(info['fechas_sync'].get(nodo, "01/01/1900 00:00:00"), 
                   format='%d/%m/%Y %H:%M:%S', errors='coerce')) 
                   for nodo in info['nodos'] if nodo not in seen and not seen.add(nodo)]
                nodos_con_fechas.sort(key=lambda x: x[1])  # Sort by timestamp
                nodos_ordenados = [nodo for nodo, _ in nodos_con_fechas]
                num_nodos = len(nodos_ordenados)
                # Obtener postes desde la fila "Nodos postes"
                postes = [nodo.split('_')[0] for nodo in nodos_ordenados]  # Ej: ['1417541', '1417542']
                # Columnas con formato Nodo_X (Poste_Y)
                columnas = [
                    'OT', 
                    'Unidad', 
                    'Cantidad Total',
                    'Fecha Sincronización',  # Esta columna ahora solo tendrá contenido para observaciones
                    *[f"Nodo_{i+1}" for i in range(num_nodos)]
                ]
                filas = []
                
                # Fila OT y Nodos (con postes reales)
                filas.append([ot, '', '', ''] + [''] * num_nodos)
                filas.append(['Nodos postes', '', '', ''] + postes)  # Mostrar postes sin sufijo
                
                # Agregar fila con fechas de sincronización de cada nodo
                fechas_fila = ['Fechas Sincronización', '', '', '']
                for nodo in nodos_ordenados:
                    fecha = info['fechas_sync'].get(nodo, "Sin fecha")
                    fechas_fila.append(fecha)
                filas.append(fechas_fila)
                
                # ===== PROCESAR CÓDIGOS N1 Y N2 =====
                def agregar_codigos(tipo_codigo, codigos_data):
                    # Diccionario para agrupar códigos por clave (potencia)
                    codigos_agrupados = {}

                    for key, nodos_data in codigos_data.items():
                        # Si esta clave no está en el diccionario, inicializarla
                        if key not in codigos_agrupados:
                            codigos_agrupados[key] = {
                                'total': 0,
                                'por_nodo': {nodo: set() for nodo in nodos_ordenados}
                            }

                        # Contar códigos por nodo y total
                        for nodo, codigos in nodos_data.items():
                            codigos_agrupados[key]['total'] += len(codigos)
                            if nodo in nodos_ordenados:
                                codigos_agrupados[key]['por_nodo'][nodo].update(codigos)

                    # Ahora, crear las filas con los códigos agrupados
                    for key, datos in codigos_agrupados.items():
                        total = datos['total']
                        # Crear la fila con suficientes elementos vacíos para todos los nodos
                        fila = [key, 'UND', total, ''] + [''] * len(nodos_ordenados)

                        # Agregar códigos por nodo
                        for i, nodo in enumerate(nodos_ordenados):
                            # Obtener los códigos para este nodo
                            codigos = datos['por_nodo'][nodo]
                            # Si hay códigos, mostrarlos separados por coma
                            if codigos:
                                fila[4 + i] = ', '.join(sorted(codigos))

                        filas.append(fila)

                if 'codigos_n1' in info:
                    agregar_codigos("N1", info['codigos_n1'])
                if 'codigos_n2' in info:
                    agregar_codigos("N2", info['codigos_n2'])
                
                # ===== MATERIALES INSTALADOS =====
                filas.append(['MATERIALES INSTALADOS', '', '', ''] + [''] * num_nodos)
                for material_key in info['materiales']:
                    try:
                        _, nombre = material_key.split('|', 1)
                        cantidades = info['materiales'][material_key]
                        total = sum(cantidades.values())
                        # Celda de fecha vacía para materiales instalados
                        fila = [nombre, 'UND', total, ''] + [cantidades.get(n, 0) for n in nodos_ordenados]
                        filas.append(fila)
                    except Exception as e:
                        logger.error(f"Error procesando material: {str(e)}")
                        continue                                
                
                # ===== MATERIALES RETIRADOS =====
                filas.append(['MATERIALES RETIRADOS', '', '', ''] + [''] * num_nodos)
                for material_key in info.get('materiales_retirados', {}):
                    try:
                        _, nombre = material_key.split('|', 1)
                        cantidades = info['materiales_retirados'][material_key]
                        total = sum(cantidades.values())
                        # Celda de fecha vacía para materiales retirados
                        fila = [nombre, 'UND', total, ''] + [cantidades.get(n, 0) for n in nodos_ordenados]
                        filas.append(fila)
                    except Exception as e:
                        logger.error(f"Error procesando material retirado: {str(e)}")
                        continue
                
                # ===== OBSERVACIONES POR NODO =====
                # Creamos un diccionario para almacenar todas las observaciones por nodo
                #observaciones_por_nodo = {nodo: [] for nodo in nodos_ordenados}
                #
                ## Recopilamos todas las observaciones de materiales y retirados
                #for mat_key in info['aspectos_materiales']:
                #    for nodo, aspectos in info['aspectos_materiales'][mat_key].items():
                #        if nodo in observaciones_por_nodo:
                #            observaciones_por_nodo[nodo].extend(aspectos)
                #
                #for mat_key in info['aspectos_retirados']:
                #    for nodo, aspectos in info['aspectos_retirados'][mat_key].items():
                #        if nodo in observaciones_por_nodo:
                #            observaciones_por_nodo[nodo].extend(aspectos)
                
                # Eliminamos duplicados y ordenamos
                #for nodo in observaciones_por_nodo:
                #    observaciones_por_nodo[nodo] = sorted(set(observaciones_por_nodo[nodo]))
                
                # Agregamos una fila para observaciones al final de los materiales
                #filas.append(['OBSERVACIONES POR NODO', '', '', ''] + [''] * num_nodos)
                #
                ## Crear una única fila para todas las observaciones, organizadas por columna de nodo
                #fila_obs = ['Observaciones', 'Obs', '', ''] + [''] * num_nodos
                #
                ## Para cada nodo, formatear sus observaciones y colocarlas en la columna correspondiente
                #for i, nodo in enumerate(nodos_ordenados):
                #    if observaciones_por_nodo[nodo]:
                #        texto_obs = '\n'.join([f"{j+1}. {obs}" for j, obs in enumerate(observaciones_por_nodo[nodo])])
                #        fila_obs[4 + i] = texto_obs
                #
                ## Solo agregar la fila si hay al menos una observación
                #if any(observaciones_por_nodo.values()):
                #    filas.append(fila_obs)
                
                # ===== OBSERVACIONES COMPLETAS =====
                filas.append(['OBSERVACIONES COMPLETAS', '', '', ''] + [''] * num_nodos)
                
                # Organizar observaciones por nodo para la sección "OBSERVACIONES COMPLETAS"
                observaciones_completas = {nodo: [] for nodo in nodos_ordenados}
                
                # Recopilar todas las observaciones con sus códigos y fechas
                for nodo in nodos_ordenados:
                    poste = nodo.split('_')[0]
                    fecha = info['fechas_sync'].get(nodo, "Sin fecha")
                    
                    # Obtener códigos asociados
                    codigos = set()
                    for key in info['codigos_n1']:
                        codigos.update(info['codigos_n1'][key].get(nodo, set()))
                    for key in info['codigos_n2']:
                        codigos.update(info['codigos_n2'][key].get(nodo, set()))
                    
                    if not codigos:
                        codigos.add("Sin código")
                    
                    # Obtener observaciones
                    aspectos = []
                    for mat_key in info['aspectos_materiales']:
                        aspectos.extend(info['aspectos_materiales'][mat_key].get(nodo, []))
                    for mat_key in info['aspectos_retirados']:
                        aspectos.extend(info['aspectos_retirados'][mat_key].get(nodo, []))
                    
                    if not aspectos:
                        aspectos.append("Sin observaciones")
                    
                    # Formatear y agregar a las observaciones completas
                    for codigo in sorted(codigos):
                        entrada = f"Código: {codigo}\n"
                        entrada += '\n'.join([f"{i+1}. {obs}" for i, obs in enumerate(aspectos)])
                        # Agregar la fecha al final de la entrada
                        entrada += f"\n(Fecha: {fecha})"
                        observaciones_completas[nodo].append(entrada)
                
                # Crear una única fila para todas las observaciones completas
                fila_obs_completas = ['Observaciones Completas', 'Obs', '', ''] + [''] * num_nodos
                
                # Para cada nodo, formatear sus observaciones completas y colocarlas en la columna correspondiente
                for i, nodo in enumerate(nodos_ordenados):
                    if observaciones_completas[nodo]:
                        # Separar cada entrada con una línea en blanco
                        texto_obs = '\n\n'.join(observaciones_completas[nodo])
                        fila_obs_completas[4 + i] = texto_obs
                
                # Agregar la fila de observaciones completas
                filas.append(fila_obs_completas)
                
                # ===== MANO DE OBRA POR NODO =====
                # Agregar encabezado para la mano de obra
                filas.append(['MANO DE OBRA POR NODO', '', '', ''] + [''] * num_nodos)
                
                # Procesar materiales instalados y retirados para calcular la mano de obra
                materiales_instalados = {}
                materiales_retirados = {}
                
                # Extraer materiales instalados
                for material_key, nodos_qty in info['materiales'].items():
                    materiales_instalados[material_key] = {}
                    for nodo, qty in nodos_qty.items():
                        if nodo in nodos_ordenados and qty > 0:
                            materiales_instalados[material_key][nodo] = float(qty)
                
                # Extraer materiales retirados
                for material_key, nodos_qty in info.get('materiales_retirados', {}).items():
                    materiales_retirados[material_key] = {}
                    for nodo, qty in nodos_qty.items():
                        if nodo in nodos_ordenados and qty > 0:
                            materiales_retirados[material_key][nodo] = float(qty)
                
                # Cargar la plantilla de mano de obra
                plantilla = cargar_plantilla_mano_obra()
                
                # Agrupar partidas por bloques para mejor visualización
                bloques = [
                    ("Postes",
                     ["APERTURA", "APLOMADA", "CONCRETADA", "HINCADA"],
                     ["POSTE"]),
                    ("Instalación luminarias",
                     ["INSTALACION LUMINARIAS"],
                     ["LUMINARIA", "FOTOCELDA", "GRILLETE", "BRAZO"]),
                    ("Conexión a tierra",
                     ["CONEXIÓN A CABLE A TIERRA", "INSTALACION KIT SPT"],
                     ["KIT DE PUESTA A TIERRA", "CONECT PERF", "CONECTOR BIME/COM", "ALAMBRE", "TUERCA", "TORNILLO"]),
                    ("Desmontaje / Transporte",
                     ["DESMONTAJE", "TRANSPORTE"],
                     ["ALAMBRE", "BRAZO", "CÓDIGO", "CABLE"])
                ]
                
                # Pre-procesamiento para unificar los tipos de desmontaje de luminarias
                partidas_unificadas = []
                tipo_desmontaje_luminarias = []
                
                for partida in plantilla:
                    descripcion = partida['DESCRIPCION MANO DE OBRA']
                    # Verificar si es una partida de desmontaje de luminarias
                    if ("DESMONTAJE" in descripcion.upper() and "LUMINARIA" in descripcion.upper() and 
                        ("CAMIONETA" in descripcion.upper() or "CANASTA" in descripcion.upper())):
                        # Guardar la descripción original para referencia
                        tipo_desmontaje_luminarias.append(descripcion)
                        # No agregar esta partida a la lista unificada aún
                    else:
                        partidas_unificadas.append(partida)
                
                # Si hay partidas de desmontaje de luminarias, crear una partida unificada
                if tipo_desmontaje_luminarias:
                    partida_unificada = {
                        'DESCRIPCION MANO DE OBRA': "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA",
                        'UNIDAD': next((p['UNIDAD'] for p in plantilla if p['DESCRIPCION MANO DE OBRA'] in tipo_desmontaje_luminarias), "UN")
                    }
                    partidas_unificadas.append(partida_unificada)
                
                # Reemplazar la plantilla original con la unificada
                plantilla_original = plantilla
                plantilla = partidas_unificadas
                
                # Crear un diccionario para almacenar la mano de obra por nodo
                mano_obra_por_nodo = {nodo: {} for nodo in nodos_ordenados}
                
                # Para cada nodo, calcular la mano de obra necesaria
                for nodo in nodos_ordenados:
                    # Para cada bloque de partidas
                    for titulo_bloque, keywords_mo, keywords_mat in bloques:
                        # Filtrar partidas del bloque actual
                        partidas_filtradas = [
                            item for item in plantilla
                            if any(kw in item['DESCRIPCION MANO DE OBRA'].upper() for kw in keywords_mo)
                        ]
                        
                        # Para cada partida, calcular la mano de obra necesaria para este nodo
                        for partida in partidas_filtradas:
                            descripcion = partida['DESCRIPCION MANO DE OBRA']
                            unidad = partida['UNIDAD']
                            
                            # Caso especial para DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA
                            cantidad_mo = 0
                            materiales_inst = []
                            materiales_ret = []
                            
                            if descripcion == "DESMONTAJE DE LUMINARIAS CANASTA/ESCALERA":
                                # Calcular la cantidad sumando todos los tipos de desmontaje
                                for desc_original in tipo_desmontaje_luminarias:
                                    cant_temp, mat_inst_temp, mat_ret_temp = calcular_cantidad_mano_obra(
                                        desc_original,
                                        materiales_instalados,
                                        materiales_retirados,
                                        nodo
                                    )
                                    cantidad_mo += cant_temp
                                    materiales_inst.extend(mat_inst_temp)
                                    materiales_ret.extend(mat_ret_temp)
                                
                                # Eliminar duplicados en las listas de materiales
                                materiales_inst = list(set(materiales_inst))
                                materiales_ret = list(set(materiales_ret))
                            else:
                                # Procesamiento normal para otras partidas
                                cantidad_mo, materiales_inst, materiales_ret = calcular_cantidad_mano_obra(
                                    descripcion,
                                    materiales_instalados,
                                    materiales_retirados,
                                    nodo
                                )
                            
                            # Solo agregar partidas con cantidad > 0
                            if cantidad_mo > 0:
                                # Si este bloque no existe en el diccionario del nodo, crearlo
                                if titulo_bloque not in mano_obra_por_nodo[nodo]:
                                    mano_obra_por_nodo[nodo][titulo_bloque] = []
                                
                                # Formatear materiales asociados
                                materiales_texto = []
                                if materiales_inst:
                                    materiales_texto.append("INST: " + ", ".join(materiales_inst))
                                if materiales_ret:
                                    materiales_texto.append("RET: " + ", ".join(materiales_ret))
                                
                                # Agregar la partida al bloque correspondiente
                                mano_obra_por_nodo[nodo][titulo_bloque].append({
                                    'descripcion': descripcion,
                                    'unidad': unidad,
                                    'cantidad': cantidad_mo,
                                    'materiales': "\n".join(materiales_texto)
                                })
                
                # Crear filas para cada bloque de mano de obra
                for titulo_bloque, _, _ in bloques:
                    # Verificar si algún nodo tiene partidas para este bloque
                    if any(titulo_bloque in mano_obra_por_nodo[nodo] for nodo in nodos_ordenados):
                        # Agregar encabezado del bloque
                        filas.append([f"BLOQUE: {titulo_bloque}", '', '', ''] + [''] * num_nodos)
                        
                        # Recopilar todas las descripciones únicas de mano de obra para este bloque
                        descripciones_unicas = set()
                        for nodo in nodos_ordenados:
                            if titulo_bloque in mano_obra_por_nodo[nodo]:
                                for partida in mano_obra_por_nodo[nodo][titulo_bloque]:
                                    descripciones_unicas.add(partida['descripcion'])
                        
                        # Para cada descripción única, crear una fila
                        for descripcion in sorted(descripciones_unicas):
                            # Encontrar la unidad (debería ser la misma para la misma descripción)
                            unidad = next((
                                partida['unidad'] 
                                for nodo in nodos_ordenados 
                                if titulo_bloque in mano_obra_por_nodo[nodo] 
                                for partida in mano_obra_por_nodo[nodo][titulo_bloque] 
                                if partida['descripcion'] == descripcion
                            ), "UND")
                            
                            # Calcular la cantidad total sumando de todos los nodos
                            cantidad_total = sum(
                                partida['cantidad']
                                for nodo in nodos_ordenados
                                if titulo_bloque in mano_obra_por_nodo[nodo]
                                for partida in mano_obra_por_nodo[nodo][titulo_bloque]
                                if partida['descripcion'] == descripcion
                            )
                            
                            # Crear la fila con la descripción, unidad y cantidad total
                            fila_mo = [descripcion, unidad, cantidad_total, ''] + [''] * num_nodos
                            
                            # Para cada nodo, agregar la información de mano de obra
                            for i, nodo in enumerate(nodos_ordenados):
                                # Buscar la partida para este nodo y descripción
                                partida_nodo = next((
                                    partida
                                    for partida in mano_obra_por_nodo[nodo].get(titulo_bloque, [])
                                    if partida['descripcion'] == descripcion
                                ), None)
                                
                                if partida_nodo:
                                    # Formatear el contenido: cantidad + materiales
                                    #contenido = f"Cantidad: {partida_nodo['cantidad']}"
                                    contenido = f"{partida_nodo['cantidad']}"
                                    #if partida_nodo['materiales']:
                                    #    contenido += f"\n\nMateriales:\n{partida_nodo['materiales']}"
                                    
                                    fila_mo[4 + i] = contenido
                            
                            # Agregar la fila a la lista de filas
                            filas.append(fila_mo)
                
                # Crear DataFrame a partir de los datos recopilados
                df = pd.DataFrame(filas, columns=columnas)
                df.to_excel(writer, sheet_name=f"OT_{ot}", index=False)

                # Acceder a la hoja creada
                sheet = writer.sheets[f"OT_{ot}"]

                # Obtener la plantilla para la mano de obra
                plantilla = cargar_plantilla_mano_obra()
                #plantiall2 = tabla_mano_obra()
                # Agregar la tabla de mano de obra en la hoja (ahora solo como referencia, ya que la mostramos en línea)
                #plantilla_mano_obra(sheet, df, plantiall2)
                # Ya no necesitamos llamar a agregar_tabla_mano_obra aquí, ya que la mano de obra se muestra en línea
                # agregar_tabla_mano_obra(sheet, df, plantilla)                           
                  
                # Cargar el archivo con openpyxl para combinar celdas
                ws = writer.sheets[f"OT_{ot}"]

                # Combinar celdas para encabezados de secciones
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):  # Comienza desde la segunda fila
                    if row[0].value in ["MATERIALES INSTALADOS", "MATERIALES RETIRADOS", "OBSERVACIONES POR NODO", 
                                       "OBSERVACIONES COMPLETAS", "MANO DE OBRA POR NODO"] or (
                                       isinstance(row[0].value, str) and row[0].value.startswith("BLOQUE: ")):
                        start_col = 1
                        end_col = len(columnas)
                        row_idx = row[0].row
                        ws.merge_cells(
                            start_row=row_idx,
                            start_column=start_col,
                            end_row=row_idx,
                            end_column=end_col
                        )
                        # Centrar el texto y aplicar formato
                        row[0].alignment = Alignment(horizontal='center', vertical='center')
                        row[0].font = Font(bold=True)
                        
                        # Aplicar color de fondo según el tipo de encabezado
                        if row[0].value == "MANO DE OBRA POR NODO":
                            row[0].fill = PatternFill("solid", fgColor="009900")  # Verde
                            row[0].font = Font(bold=True, color="FFFFFF")  # Texto blanco
                        elif isinstance(row[0].value, str) and row[0].value.startswith("BLOQUE: "):
                            row[0].fill = PatternFill("solid", fgColor="AAAAAA")  # Gris
                        else:
                            row[0].fill = PatternFill("solid", fgColor="DDDDDD")  # Gris claro
                
                # Configurar el ajuste de texto para las celdas de observaciones y mano de obra
                for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                    for cell in row:
                        if isinstance(cell.value, str) and '\n' in cell.value:
                            cell.alignment = Alignment(wrap_text=True, vertical='top')
                            # Ajustar altura de fila para acomodar el texto
                            if cell.row > 1:  # Evitar ajustar la fila de encabezados
                                ws.row_dimensions[cell.row].height = max(15 * cell.value.count('\n') + 15, 15)
              
            # Eliminar hoja temporal si se crearon hojas
            if sheets_created:
                writer.book.remove(writer.book[temp_sheet_name])
        
        
        # Manejo de errores
        if not sheets_created:
            writer.book.remove(writer.book[temp_sheet_name])
            df_error = pd.DataFrame({
                'Error': [
                    'Datos no válidos - Razones posibles:',
                    '1. Columnas requeridas faltantes',
                    '2. Valores "NINGUNO" o 0 en todos los registros',
                    '3. Formato de archivo incorrecto'
                ]
            })
            df_error.to_excel(writer, sheet_name='Errores', index=False)

    output.seek(0)
    return output

@app.post("/upload/")
async def subir_archivos(
    files: list[UploadFile] = File(...),
    tipo_archivo: str = Form(..., description="Tipo de archivo: modernizacion o mantenimiento")
):
    try:
        datos_combinados = defaultdict(lambda: {
            'nodos': set(),
            'codigos_n1': defaultdict(lambda: defaultdict(set)),
            'codigos_n2': defaultdict(lambda: defaultdict(set)),
            'materiales': defaultdict(lambda: defaultdict(int)),
            'materiales_retirados': defaultdict(lambda: defaultdict(int)),
            'aspectos_materiales': defaultdict(lambda: defaultdict(lambda: set())),
            'aspectos_retirados': defaultdict(lambda: defaultdict(lambda: set())),
            'fechas_sync': defaultdict(str)
        })
        
        datos_por_barrio_combinados = defaultdict(lambda: {
            'materiales_instalados': defaultdict(lambda: defaultdict(int)),
            'materiales_retirados': defaultdict(lambda: defaultdict(int)),
        })

        dfs_originales_combinados = {}  # Inicializar diccionario para DataFrames originales

        for file in files:
            try:
                if tipo_archivo == 'modernizacion':
                    # Procesar archivo y obtener datos, datos_por_barrio y DataFrames originales
                    datos, datos_por_barrio, dfs_originales = procesar_archivo_modernizacion(file)
                    # Combinar DataFrames originales
                    dfs_originales_combinados.update(dfs_originales)
                elif tipo_archivo == 'mantenimiento':
                    datos = procesar_archivo_mantenimiento(file)
                else:
                    raise ValueError("Tipo de archivo no válido")

                # Combinar datos por barrio (solo para modernización)
                if tipo_archivo == 'modernizacion':
                    for barrio, barrio_info in datos_por_barrio.items():
                        # Materiales instalados
                        for mat_key, ot_counts in barrio_info['materiales_instalados'].items():
                            for ot, cantidad in ot_counts.items():
                                datos_por_barrio_combinados[barrio]['materiales_instalados'][mat_key][ot] += cantidad
                        # Materiales retirados
                        for mat_key, ot_counts in barrio_info['materiales_retirados'].items():
                            for ot, cantidad in ot_counts.items():
                                datos_por_barrio_combinados[barrio]['materiales_retirados'][mat_key][ot] += cantidad

                # Combinar datos comunes
                for ot, info in datos.items():
                    # Actualizar nodos y fechas
                    datos_combinados[ot]['nodos'].update(info['nodos'])
                    for nodo, fecha in info.get('fechas_sync', {}).items():
                        datos_combinados[ot]['fechas_sync'][nodo] = fecha
                    
                    # Combinar materiales instalados
                    for mat_key, cantidades in info.get('materiales', {}).items():
                        for nodo, cantidad in cantidades.items():
                            datos_combinados[ot]['materiales'][mat_key][nodo] += cantidad
                    
                    # Procesar códigos y materiales retirados solo para modernización
                    if tipo_archivo == 'modernizacion':
                        # Códigos N1 y N2
                        for key, nodos_data in info.get('codigos_n1', {}).items():
                            for nodo, codigos in nodos_data.items():
                                datos_combinados[ot]['codigos_n1'][key][nodo].update(codigos)
                        for key, nodos_data in info.get('codigos_n2', {}).items():
                            for nodo, codigos in nodos_data.items():
                                datos_combinados[ot]['codigos_n2'][key][nodo].update(codigos)
                        
                        # Materiales retirados y aspectos
                        for mat_key, cantidades in info.get('materiales_retirados', {}).items():
                            for nodo, cantidad in cantidades.items():
                                datos_combinados[ot]['materiales_retirados'][mat_key][nodo] += cantidad
                        for mat_key, nodos_aspectos in info.get('aspectos_materiales', {}).items():
                            for nodo, aspectos in nodos_aspectos.items():
                                datos_combinados[ot]['aspectos_materiales'][mat_key][nodo].update(aspectos)
                        for mat_key, nodos_aspectos in info.get('aspectos_retirados', {}).items():
                            for nodo, aspectos in nodos_aspectos.items():
                                datos_combinados[ot]['aspectos_retirados'][mat_key][nodo].update(aspectos)

            except Exception as e:
                logger.error(f"Error con {file.filename}: {str(e)}")
                continue

        # Generar el Excel con todos los datos combinados
        excel_final = generar_excel(datos_combinados, datos_por_barrio_combinados, dfs_originales_combinados)
        
        return Response(
            content=excel_final.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=resultado.xlsx"}
        )

    except Exception as e:
        logger.critical(f"Error global: {str(e)}")
        raise HTTPException(500, detail=str(e))
    
      
    