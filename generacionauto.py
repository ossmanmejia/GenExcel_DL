import pandas as pd

# Cargar el archivo Excel con las tres hojas
archivo = 'postulaciones_actualizado.xlsx'
hoja_inmersion = pd.read_excel(archivo, sheet_name='Inmersiones')
hoja_taller = pd.read_excel(archivo, sheet_name='Talleres')
hoja_evaluadores = pd.read_excel(archivo, sheet_name='Evaluadores')

# Crear un diccionario para almacenar las evaluaciones de cada evaluador
evaluaciones_por_evaluador = {}

# Recorrer la hoja de Evaluadores para obtener las propuestas asignadas a cada evaluador
for index, row in hoja_evaluadores.iterrows():
    codigo_evaluador = row['Código']

    # Obtener todas las evaluaciones asignadas
    evaluaciones = row[['Evaluación 1', 'Evaluación 2', 'Evaluación 3',
                        'Evaluación 4', 'Evaluación 5', 'Evaluación 6',
                        'Evaluación 7', 'Evaluación 8', 'Evaluación 9',
                        'Evaluación 10']].dropna().tolist()

    if evaluaciones:
        evaluaciones_por_evaluador[codigo_evaluador] = evaluaciones

# Generar un archivo Excel para cada evaluador con sus inmersiones y talleres asignados
for codigo_evaluador, codigos_evaluaciones in evaluaciones_por_evaluador.items():
    # Filtrar inmersiones y talleres asignados
    inmersiones_asignadas = hoja_inmersion[hoja_inmersion['Código'].isin(
        codigos_evaluaciones)]
    talleres_asignados = hoja_taller[hoja_taller['Código'].isin(
        codigos_evaluaciones)]

    # Eliminar las columnas Evaluador 1, Evaluador 2 y Evaluador 3, si existen
    inmersiones_asignadas = inmersiones_asignadas.drop(
        columns=['Evaluador 1', 'Evaluador 2', 'Evaluador 3'], errors='ignore')
    talleres_asignados = talleres_asignados.drop(
        columns=['Evaluador 1', 'Evaluador 2', 'Evaluador 3'], errors='ignore')

    # Crear un archivo Excel solo si hay inmersiones o talleres asignados
    if not inmersiones_asignadas.empty or not talleres_asignados.empty:
        # Aquí se usa solo el código del evaluador
        nombre_archivo = f'{codigo_evaluador}.xlsx'

        with pd.ExcelWriter(nombre_archivo) as writer:
            if not inmersiones_asignadas.empty:
                inmersiones_asignadas.to_excel(
                    writer, sheet_name='Inmersiones', index=False)
            if not talleres_asignados.empty:
                talleres_asignados.to_excel(
                    writer, sheet_name='Talleres', index=False)

print("Archivos generados correctamente con el código de cada evaluador.")
