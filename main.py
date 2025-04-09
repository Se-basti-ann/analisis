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
                                     
                original_nodo = str(fila["1.NODO DEL POSTE."]).strip()  # Limpiar y convertir a string
                
                # Guardar la fecha de sincronización para este nodo/registro
                fecha_sincronizacion = fila["FechaSincronizacion"]
    
                # Si el nodo es 0, asignar un identificador único por OT
                if original_nodo in ['0', '0.0']:
                    counter_0[ot] += 1
                    nodo = f"0_{counter_0[ot]}"  # Ejemplo: 0_1, 0_2, etc.
                else:
                    count = datos[ot]['nodo_counts'][original_nodo] + 1
                    datos[ot]['nodo_counts'][original_nodo] = count
                    nodo = f"{original_nodo}_{count}" if count > 1 else original_nodo                  
                #datos[ot]['nodos'].add(nodo)             
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
                    
                    # Process N1 and N2 code information separately - don't add materials here
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
                
                    if (
                        pd.notna(codigo_n2)
                        and str(codigo_n2).strip() not in ['', '0', '0.0']
                        or pd.notna(potencia_n2) != "0"
                        and float(potencia_n2) != 0   
                        
                        ):
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
                                datos[ot]['aspectos_retirados'][key][nodo].add(aspecto_limpio)                               
                                               
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
            
        return df.to_dict('records')
        
    except Exception as e:
        logger.error(f"Error cargando plantilla: {str(e)}")
        return []

def agregar_tabla_mano_obra(worksheet, df, plantilla):

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
        worksheet.cell(row=row_num, column=1, value=item['DESCRIPCION MANO DE OBRA']).border = cell_border
        worksheet.cell(row=row_num, column=2, value=item['UNIDAD']).border = cell_border
        worksheet.cell(row=row_num, column=3, value=item['CANTIDAD']).border = cell_border
        #worksheet.cell(row=row_num, column=4, value=item['VALOR UNITARIO']).number_format = '"$ "#,##0.00'
        #worksheet.cell(row=row_num, column=5, value=f'=C{row_num}*D{row_num}').number_format = '"$ "#,##0.00'
    
    # Ajustar anchos de columna
    worksheet.column_dimensions['A'].width = 45
    worksheet.column_dimensions['B'].width = 10
    worksheet.column_dimensions['C'].width = 12
    #worksheet.column_dimensions['D'].width = 15
    #worksheet.column_dimensions['E'].width = 15

def generar_excel(datos_combinados, datos_por_barrio_combinados, dfs_originales_combinados):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        temp_sheet_name = "PlantillaInicial"
        writer.book.create_sheet(temp_sheet_name)
        sheets_created = False        
        
        generate_resumen_general(writer, datos_combinados)
        
        generate_resumen_tecnicos(writer, datos_combinados, dfs_originales_combinados)  
                         
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
                observaciones_por_nodo = {nodo: [] for nodo in nodos_ordenados}
                
                # Recopilamos todas las observaciones de materiales y retirados
                for mat_key in info['aspectos_materiales']:
                    for nodo, aspectos in info['aspectos_materiales'][mat_key].items():
                        if nodo in observaciones_por_nodo:
                            observaciones_por_nodo[nodo].extend(aspectos)
                
                for mat_key in info['aspectos_retirados']:
                    for nodo, aspectos in info['aspectos_retirados'][mat_key].items():
                        if nodo in observaciones_por_nodo:
                            observaciones_por_nodo[nodo].extend(aspectos)
                
                # Eliminamos duplicados y ordenamos
                for nodo in observaciones_por_nodo:
                    observaciones_por_nodo[nodo] = sorted(set(observaciones_por_nodo[nodo]))
                
                # Agregamos una fila para observaciones al final de los materiales
                filas.append(['OBSERVACIONES POR NODO', '', '', ''] + [''] * num_nodos)
                
                # Para cada nodo que tenga observaciones, agregamos una fila
                for nodo in nodos_ordenados:
                    if observaciones_por_nodo[nodo]:
                        poste = nodo.split('_')[0]
                        texto_obs = '\n'.join([f"{i+1}. {obs}" for i, obs in enumerate(observaciones_por_nodo[nodo])])
                        fecha = info['fechas_sync'].get(nodo, "Sin fecha")
                        
                        # Creamos una fila con observaciones solo para este nodo
                        fila = [f"Obs. {poste}", 'Obs', '', fecha] + [''] * num_nodos
                        # Colocamos las observaciones en la columna correcta
                        nodo_idx = nodos_ordenados.index(nodo)
                        fila[4 + nodo_idx] = texto_obs
                        filas.append(fila)
                
                # ===== GARANTIZAR COBERTURA TOTAL DE NODOS =====
                filas.append(['OBSERVACIONES COMPLETAS', '', '', ''] + [''] * num_nodos)

                # Diccionario para trackear nodos procesados
                observaciones_ordenadas = []

                # Recorrer nodos en ORDEN CRONOLÓGICO ORIGINAL
                for nodo in nodos_ordenados:
                    poste = nodo.split('_')[0]

                    # 1. Obtener códigos asociados
                    codigos = set()
                    for key in info['codigos_n1']:
                        codigos.update(info['codigos_n1'][key].get(nodo, set()))
                    for key in info['codigos_n2']:
                        codigos.update(info['codigos_n2'][key].get(nodo, set()))

                    if not codigos:
                        codigos.add("Sin código")

                    # 2. Obtener observaciones
                    aspectos = []
                    for mat_key in info['aspectos_materiales']:
                        aspectos.extend(info['aspectos_materiales'][mat_key].get(nodo, []))
                    for mat_key in info['aspectos_retirados']:
                        aspectos.extend(info['aspectos_retirados'][mat_key].get(nodo, []))

                    if not aspectos:
                        aspectos.append("Sin observaciones")

                    # 3. Registrar todas las entradas
                    for codigo in sorted(codigos):
                        texto_obs = '\n'.join([f"{i+1}. {obs}" for i, obs in enumerate(aspectos)])
                        fecha = info['fechas_sync'].get(nodo, "Sin fecha")
                        observaciones_ordenadas.append({
                            'orden': len(observaciones_ordenadas) + 1,
                            'clave': f"{poste} - {codigo}",
                            'observaciones': texto_obs,
                            'nodo': nodo,
                            'fecha': fecha  # Guardar la fecha para ordenar y mostrar
                        })

                # 4. Agregar nodos faltantes
                nodos_procesados = {obs['nodo'] for obs in observaciones_ordenadas}
                for nodo in nodos_ordenados:
                    if nodo not in nodos_procesados:
                        poste = nodo.split('_')[0]
                        fecha = info['fechas_sync'].get(nodo, "Sin fecha")
                        observaciones_ordenadas.append({
                            'orden': len(observaciones_ordenadas) + 1,
                            'clave': f"{poste} - Sin código",
                            'observaciones': "1. Sin observaciones",
                            'nodo': nodo,
                            'fecha': fecha
                        })

                # 5. Ordenar por fecha de sincronización (de menor a mayor)
                observaciones_ordenadas_por_fecha = sorted(
                    observaciones_ordenadas,
                    key=lambda x: pd.to_datetime(x['fecha'], errors='coerce', format='%d/%m/%Y %H:%M:%S')
                )

                # 6. Agregar al Excel - SOLO para observaciones incluimos la fecha
                for obs in observaciones_ordenadas_por_fecha:
                    fila = [
                        obs['clave'],
                        'Obs',
                        obs['observaciones'],
                        obs['fecha'],  # Incluir la fecha SOLO para las observaciones
                        *['' for _ in range(num_nodos)]
                    ]
                    filas.append(fila)
                
                # Crear DataFrame a partir de los datos recopilados
                df = pd.DataFrame(filas, columns=columnas)
                df.to_excel(writer, sheet_name=f"OT_{ot}", index=False)

                # Acceder a la hoja creada
                sheet = writer.sheets[f"OT_{ot}"]

                # Obtener la plantilla para la mano de obra
                plantilla = cargar_plantilla_mano_obra()

                # Agregar la tabla de mano de obra en la hoja
                agregar_tabla_mano_obra(sheet, df, plantilla)                                
                # Cargar el archivo con openpyxl para combinar celdas
                ws = writer.sheets[f"OT_{ot}"]

                # Combinar celdas para "MATERIALES INSTALADOS" y "MATERIALES RETIRADOS"
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):  # Comienza desde la segunda fila
                    if row[0].value in ["MATERIALES INSTALADOS", "MATERIALES RETIRADOS", "OBSERVACIONES POR NODO"]:
                        start_col = 1
                        end_col = len(columnas)
                        row_idx = row[0].row
                        ws.merge_cells(
                            start_row=row_idx,
                            start_column=start_col,
                            end_row=row_idx,
                            end_column=end_col
                        )
                    # Opcional: Centrar el texto
                    row[0].alignment = row[0].alignment.copy(horizontal='center', vertical='center')                

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
    
      
    