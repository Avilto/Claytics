import sys
import json
import pandas as pd

def determine_sector(columns):
    # Diccionarios de palabras clave por sector
    sectors = {
        "Retail / Ventas": ["precio", "stock", "venta", "producto", "cliente", "sku", "descuento", "factura", "categoria"],
        "Salud": ["paciente", "diagnostico", "tratamiento", "medico", "sintoma", "historial", "clinica", "dosis"],
        "Finanzas": ["cuenta", "saldo", "transaccion", "credito", "interes", "banco", "inversion", "prestamo", "riesgo"],
        "Educacion": ["alumno", "estudiante", "profesor", "calificacion", "nota", "curso", "semestre", "matricula"],
        "Marketing": ["campaña", "clics", "impresiones", "conversion", "audiencia", "engagement", "leads"]
    }
    
    col_str = " ".join(columns).lower()
    
    sector_scores = {sector: 0 for sector in sectors}
    
    for sector, keywords in sectors.items():
        for keyword in keywords:
            if keyword in col_str:
                sector_scores[sector] += 1
                
    best_sector = max(sector_scores, key=sector_scores.get)
    
    if sector_scores[best_sector] == 0:
        return "General / Otro"
        
    return best_sector

def analyze_file(file_path):
    try:
        # Detectar extensión
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path, nrows=0) # Solo lee cabeceras
        elif file_path.lower().endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path, nrows=0)
        else:
            return {"error": "Formato no soportado"}
            
        columns = df.columns.tolist()
        sector = determine_sector(columns)
        
        return {
            "status": "success",
            "file": file_path,
            "columns_detected": len(columns),
            "predicted_sector": sector
        }
    except Exception as e:
        return {"error": str(e)}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Por favor provee la ruta del archivo. Ej: python data_profiler.py archivo.csv"}))
        sys.exit(1)
        
    file_path = sys.argv[1]
    result = analyze_file(file_path)
    print(json.dumps(result, indent=2))
