from email import header
import os
from openpyxl.utils import column_index_from_string
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
            'VALOR UNITARIO',
            'VALOR TOTAL'
        ]
        
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Plantilla no tiene las columnas requeridas")
            
        return df.to_dict('records')
        
    except Exception as e:
        logger.error(f"Error cargando plantilla: {str(e)}")
        return []
           

def procesar_archivo_modernizacion(file: UploadFile):
    try:
        contenido = file.file.read()
        xls = pd.ExcelFile(BytesIO(contenido))
        
        datos = defaultdict(lambda: {
            'nodos': set(),
            'codigos_n1': defaultdict(lambda: defaultdict(set)),
            'codigos_n2': defaultdict(lambda: defaultdict(set)),
            'materiales': defaultdict(lambda: defaultdict(int)),    
            'materiales_retirados': defaultdict(lambda: defaultdict(int))        
        })                 
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
                "1. Describa Aspectos que Considere se deben tener en cuenta."
            }
            if not required_columns.issubset(df.columns):
                continue          
            
            # PROCESAR MATERIALES RETIRADOS
            pattern_codigo = re.compile(r'^\d+\.CODIGO DE (LUMINARIA|BOMBILLA|FOTOCELDA) RETIRADA (N\d+)\.?$', re.IGNORECASE)
            pattern_potencia = re.compile(r'^\d+\.POTENCIA DE (LUMINARIA|BOMBILLA) RETIRADA (N\d+)\.\(W\)$', re.IGNORECASE)
            
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
                nodo = str(fila["1.NODO DEL POSTE."])
                datos[ot]['nodos'].add(nodo)                                
                
                codigo_n1 = fila["2.CODIGO DE LUMINARIA INSTALADA N1."]
                potencia_n1 = fila["3.POTENCIA DE LUMINARIA INSTALADA (W)"]
                
                for col in columnas_bh_bo:
                    cantidad = fila[col]
                    
                    if pd.notna(cantidad) and float(cantidad) > 0:
                        nombre_material = str(col).split('.', 1)[-1].strip().upper()
                        key = f"MATERIAL_RETIRADO|{nombre_material}"
                        datos[ot]['materiales_retirados'][key][nodo] += cantidad
                    
                    #if pd.notna(cantidad) and float(cantidad) > 0:
                    #    nombre_material = str(col).split('.', 1)[-1].strip().upper()
                    #    key = f"MATERIAL|{nombre_material}"
                    #    datos[ot]['materiales'][key][nodo] += cantidad
                        
                if pd.notna(codigo_n1) and pd.notna(potencia_n1):
                    key = f"CODIGO 1 LUMINARIA INSTALADA {potencia_n1} W"
                    datos[ot]['codigos_n1'][key][nodo].add(str(codigo_n1).strip().upper())
                    
                codigo_n2 = fila["6.CODIGO DE LUMINARIA INSTALADA N2."]
                potencia_n2 = fila["7.POTENCIA DE LUMINARIA INSTALADA (W)"]
                
                if pd.notna(codigo_n2) and pd.notna(potencia_n2):
                    key = f"CODIGO 2 LUMINARIA INSTALADA {potencia_n2} W"
                    datos[ot]['codigos_n2'][key][nodo].add(str(codigo_n2).strip().upper())
                    
                # PROCESAR MATERIALES RETIRADOS
                for (tipo, n), col_codigo in codigo_columns.items():
                    codigo_val = fila.get(col_codigo)
                    if pd.notna(codigo_val) and str(codigo_val).strip() != '':
                        potencia_val = None
                        if tipo in ['LUMINARIA', 'BOMBILLA']:
                            col_potencia = potencia_columns.get((tipo, n), None)
                            if col_potencia:
                                potencia_val = fila.get(col_potencia)
                                
                            # Construir nombre del material
                            if tipo == 'FOTOCELDA':
                                entry_name = f"FOTOCELDA RETIRADA {n}"
                            else:
                                potencia_str = f"{potencia_val}W" if pd.notna(potencia_val) else ''
                                entry_name = f"{tipo} RETIRADA {n} {potencia_str}".strip()
                            
                            key = f"MATERIAL_RETIRADO|{entry_name}"
                            datos[ot]['materiales_retirados'][key][nodo] += 1
                # Procesar materiales INSTALADOS
                for mat_col, cant_col in zip(material_cols, cantidad_cols):
                    material = fila[mat_col]
                    cantidad = fila[cant_col]
                    
                    if pd.notna(material) and pd.notna(cantidad):
                        if str(material).strip().upper() not in ['NINGUNO', 'SIN DATOS', 'NA']:
                            if float(cantidad) > 0.0:
                                key = f"MATERIAL|{material}".strip().upper()
                                datos[ot]['materiales'][key][nodo] += cantidad        
                                               
        return datos

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

        for hoja in xls.sheet_names:
            df = xls.parse(hoja).rename(columns=lambda x: str(x).strip())
            
            required_columns = {"6.Nro.Orden Energis", "5.Nodo"}
            if not required_columns.issubset(df.columns):
                logger.warning(f"Hoja {hoja} omitida - Columnas faltantes")
                continue

            # Identificar columnas MATERIAL y CANTIDAD exactas
            material_cols = [col for col in df.columns if re.match(r'^MATERIAL\s\d+$', col, re.IGNORECASE)]
            cantidad_cols = [col for col in df.columns if re.match(r'^CANTIDAD MATERIAL\s\d+$', col, re.IGNORECASE)]
            
            # Emparejar columnas por número
            materiales = {}
            for col in material_cols:
                num = re.search(r'\d+', col).group()
                materiales[num] = {'material': col}
            
            for col in cantidad_cols:
                num = re.search(r'\d+', col).group()
                if num in materiales:
                    materiales[num]['cantidad'] = col

            for _, fila in df.iterrows():
                ot = fila["6.Nro.Orden Energis"]
                nodo = str(fila["5.Nodo"])
                datos[ot]['nodos'].add(nodo)

                # Procesar cada par material-cantidad
                for num, cols in materiales.items():
                    if 'cantidad' not in cols:
                        continue
                    
                    material = fila[cols['material']]
                    cantidad = fila[cols['cantidad']]
                    
                    try:
                        # Validar material
                        if pd.isna(material) or str(material).strip().upper() in ['NINGUNO', 'NA', '']:
                            continue
                            
                        # Convertir cantidad a número
                        cantidad_val = float(cantidad) if not isinstance(cantidad, pd.Timestamp) else 0.0
                        if cantidad_val <= 0:
                            continue
                            
                        # Registrar material
                        key = f"MATERIAL|{str(material).strip().upper()}"
                        datos[ot]['materiales'][key][nodo] += cantidad_val
                        
                    except Exception as e:
                        logger.error(f"Error procesando fila {_}: {str(e)}")
                        continue

        return datos

    except Exception as e:
        logger.error(f"Error procesando {file.filename}: {str(e)}")
        raise HTTPException(500, detail=f"Error en archivo {file.filename}")

def generar_excel(datos):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        temp_sheet_name = "PlantillaInicial"
        writer.book.create_sheet(temp_sheet_name)
        sheets_created = False
        
        if datos:                        
            
            for ot, info in datos.items():
                try:
                    nodos_ordenados = sorted(info['nodos'], key=lambda x: int(x) if x.isdigit() else x)
                except ValueError:
                    nodos_ordenados = sorted(info['nodos'])
                
                columnas = ['OT', 'Unidad', 'Cantidad Total'] + [f"Nodo_{i+1}" for i in range(len(nodos_ordenados))]
                filas = []
                
                # Fila OT
                filas.append([ot, '', ''] + [''] * len(nodos_ordenados))
                # Fila Nodos
                filas.append(['Nodos postes', '', ''] + nodos_ordenados)
                
                # Procesar Códigos N1 (solo para Modernización)
                if 'codigos_n1' in info:
                    for key, nodos_data in info['codigos_n1'].items():
                        total = sum(len(codigos) for codigos in nodos_data.values())
                        fila = [key, 'UND', total]
                        for nodo in nodos_ordenados:
                            codigos = ', '.join(nodos_data.get(nodo, set()))
                            fila.append(codigos if codigos else '')
                        filas.append(fila)
                
                # Procesar Códigos N2 (solo para Modernización)
                if 'codigos_n2' in info:
                    for key, nodos_data in info['codigos_n2'].items():
                        total = sum(len(codigos) for codigos in nodos_data.values())
                        fila = [key, 'UND', total]
                        for nodo in nodos_ordenados:
                            codigos = ', '.join(nodos_data.get(nodo, set()))
                            fila.append(codigos if codigos else '')
                        filas.append(fila)
                        
                filas.append(['MATERIALES INSTALADOS', '', ''] + [''] * len(nodos_ordenados)) #"", ""] + [''] * len(nodos_ordenados))
                
                # Procesar Materiales Instalados (orden alfabético)
                for material_key in sorted(info['materiales'].keys(), key=lambda x: x.split('|', 1)[1].lower()):
                    try:
                        _, nombre = material_key.split('|', 1)
                        cantidades = info['materiales'][material_key]
                        total = sum(cantidades.values())
                        fila = [
                            nombre,
                            'UND',
                            total,
                            *[int(cant) if isinstance(cant, (int, float)) else 0 for cant in (cantidades.get(nodo, 0) for nodo in nodos_ordenados)]
                        ]
                        filas.append(fila)
                    except Exception as e:
                        logger.error(f"Error procesando material: {str(e)}")
                        continue

                # Agregar encabezado de Materiales Retirados
                filas.append(['MATERIALES RETIRADOS', '', ''] + [''] * len(nodos_ordenados))

                # Procesar Materiales Retirados (orden alfabético)
                for material_key in sorted(info.get('materiales_retirados', {}).keys(), key=lambda x: x.split('|', 1)[1].lower()):
                    try:
                        _, nombre = material_key.split('|', 1)
                        cantidades = info['materiales_retirados'][material_key]
                        total = sum(cantidades.values())
                        fila = [
                            nombre,
                            'UND',
                            total,
                            *[int(cant) if isinstance(cant, (int, float)) else 0 for cant in (cantidades.get(nodo, 0) for nodo in nodos_ordenados)]
                        ]
                        filas.append(fila)
                    except Exception as e:
                        logger.error(f"Error procesando material retirado: {str(e)}")
                        continue
                
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
                    if row[0].value in ["MATERIALES INSTALADOS", "MATERIALES RETIRADOS"]:
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


def agregar_tabla_mano_obra(worksheet, df, plantilla):
    from openpyxl.styles import Font, Border, Side, Alignment
    
    # Configurar estilos
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
    headers = ['DESCRIPCION MANO DE OBRA', 'UNIDAD', 'CANTIDAD', 'VALOR UNITARIO', 'VALOR TOTAL']
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
        worksheet.cell(row=row_num, column=4, value=item['VALOR UNITARIO']).number_format = '"$ "#,##0.00'
        worksheet.cell(row=row_num, column=5, value=f'=C{row_num}*D{row_num}').number_format = '"$ "#,##0.00'
    
    # Ajustar anchos de columna
    worksheet.column_dimensions['A'].width = 45
    worksheet.column_dimensions['B'].width = 10
    worksheet.column_dimensions['C'].width = 12
    worksheet.column_dimensions['D'].width = 15
    worksheet.column_dimensions['E'].width = 15

@app.post("/upload/")
async def subir_archivos(
    files: list[UploadFile] = File(...),
    tipo_archivo: str = Form(..., description="Tipo de archivo: modernizacion o mantenimiento")  # 'modernizacion' o 'mantenimiento'
):
    try:
        datos_combinados = defaultdict(lambda: {
            'nodos': set(),
            'codigos_n1': defaultdict(lambda: defaultdict(set)),
            'codigos_n2': defaultdict(lambda: defaultdict(set)),
            'materiales': defaultdict(lambda: defaultdict(int)),
            'materiales_retirados': defaultdict(lambda: defaultdict(int))
        })

        for file in files:
            try:
                if tipo_archivo == 'modernizacion':
                    datos = procesar_archivo_modernizacion(file)
                elif tipo_archivo == 'mantenimiento':
                    datos = procesar_archivo_mantenimiento(file)
                else:
                    raise ValueError("Tipo de archivo no válido")

                # Combinar datos comunes
                for ot, info in datos.items():
                    # Combinar nodos
                    datos_combinados[ot]['nodos'].update(info['nodos'])
                    
                    # Materiales instalados
                    for mat_key, cantidades in info.get('materiales', {}).items():
                        for nodo, cantidad in cantidades.items():
                            datos_combinados[ot]['materiales'][mat_key][nodo] += cantidad                                        
                    
                    # Combinar códigos solo para modernización
                    if tipo_archivo == 'modernizacion':
                        # Códigos N1
                        for key, nodos_data in info.get('codigos_n1', {}).items():
                            for nodo, codigos in nodos_data.items():
                                datos_combinados[ot]['codigos_n1'][key][str(nodo).upper()].update(codigos)
                        
                        # Códigos N2
                        for key, nodos_data in info.get('codigos_n2', {}).items():
                            for nodo, codigos in nodos_data.items():
                                datos_combinados[ot]['codigos_n2'][key][str(nodo).upper()].update(codigos)
                    
                        # Materiales retirados
                        for mat_key, cantidades in info.get('materiales_retirados', {}).items():
                            for nodo, cantidad in cantidades.items():
                                datos_combinados[ot]['materiales_retirados'][mat_key][nodo] += cantidad

            except Exception as e:
                logger.error(f"Error con {file.filename}: {str(e)}")
                continue

        excel_final = generar_excel(datos_combinados)
        
        return Response(
            content=excel_final.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=resultado.xlsx"}
        )

    except Exception as e:
        logger.critical(f"Error global: {str(e)}")
        raise HTTPException(500, detail=str(e))